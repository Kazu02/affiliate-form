/**
 * 顧客向けLINE基盤から既存GASへ渡す署名付きブリッジの検証。
 * 秘密値はソースに書かず、CUSTOMER_LINE_BRIDGE_SECRET を Script Properties に保存する。
 */

const CUSTOMER_LINE_BRIDGE_SECRET_PROPERTY = 'CUSTOMER_LINE_BRIDGE_SECRET';
const DIRECT_LINE_WEBHOOK_ENABLED_PROPERTY = 'LINE_DIRECT_WEBHOOK_ENABLED';
const BRIDGE_NONCE_PREFIX = 'CUSTOMER_LINE_BRIDGE_NONCE_';
const BRIDGE_EVENT_PREFIX = 'CUSTOMER_LINE_BRIDGE_EVENT_';

function _customerLineBridgeResponse(accepted, code) {
  return ContentService.createTextOutput(JSON.stringify({
    status: accepted ? 'ok' : 'rejected',
    bridgeAccepted: accepted === true,
    code: code || (accepted ? 'accepted' : 'rejected')
  })).setMimeType(ContentService.MimeType.JSON);
}

function _customerLineBridgeConstantTimeEqual(left, right) {
  left = String(left || '');
  right = String(right || '');
  var mismatch = left.length ^ right.length;
  var length = Math.max(left.length, right.length);
  for (var index = 0; index < length; index++) {
    mismatch |= (left.charCodeAt(index % Math.max(1, left.length)) || 0)
      ^ (right.charCodeAt(index % Math.max(1, right.length)) || 0);
  }
  return mismatch === 0;
}

function _customerLineBridgeNonceDigest(nonce) {
  var bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    nonce,
    Utilities.Charset.UTF_8
  );
  return bytes.map(function(byte) {
    var value = byte < 0 ? byte + 256 : byte;
    return ('0' + value.toString(16)).slice(-2);
  }).join('');
}

function _customerLineBridgeReserveNonce(nonce, nowMillis) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return false;
  try {
    var props = PropertiesService.getScriptProperties();
    var key = BRIDGE_NONCE_PREFIX + _customerLineBridgeNonceDigest(nonce);
    var existingExpiry = Number(props.getProperty(key) || 0);
    if (existingExpiry > nowMillis) return false;

    var all = props.getProperties();
    Object.keys(all).forEach(function(candidate) {
      if (candidate.indexOf(BRIDGE_NONCE_PREFIX) !== 0) return;
      if (Number(all[candidate] || 0) <= nowMillis) props.deleteProperty(candidate);
    });
    props.setProperty(key, String(nowMillis + 10 * 60 * 1000));
    return true;
  } finally {
    lock.releaseLock();
  }
}

function _customerLineBridgeClaimEvent(idempotencyKey, nowMillis) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return { accepted: false, duplicate: false };
  try {
    var props = PropertiesService.getScriptProperties();
    var key = BRIDGE_EVENT_PREFIX + _customerLineBridgeNonceDigest(idempotencyKey);
    var existing = String(props.getProperty(key) || '').split(':');
    var existingExpiry = Number(existing[1] || 0);
    if (existingExpiry > nowMillis) {
      return { accepted: existing[0] === 'done', duplicate: existing[0] === 'done' };
    }

    var all = props.getProperties();
    Object.keys(all).forEach(function(candidate) {
      if (candidate.indexOf(BRIDGE_EVENT_PREFIX) !== 0) return;
      var expiry = Number(String(all[candidate] || '').split(':')[1] || 0);
      if (expiry <= nowMillis) props.deleteProperty(candidate);
    });
    props.setProperty(key, 'processing:' + String(nowMillis + 2 * 60 * 1000));
    return { accepted: true, duplicate: false };
  } finally {
    lock.releaseLock();
  }
}

function _customerLineBridgeFinishEvent(idempotencyKey, succeeded) {
  if (!idempotencyKey) return;
  var lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    var props = PropertiesService.getScriptProperties();
    var key = BRIDGE_EVENT_PREFIX + _customerLineBridgeNonceDigest(idempotencyKey);
    if (succeeded) {
      props.setProperty(key, 'done:' + String(Date.now() + 7 * 24 * 60 * 60 * 1000));
    } else {
      props.deleteProperty(key);
    }
  } finally {
    lock.releaseLock();
  }
}

function _verifyCustomerLineBridgeRequest(e, rawBody, parsedBody) {
  var isEnvelope = parsedBody
    && parsedBody.schemaVersion === 1
    && typeof parsedBody.organizationId === 'string'
    && typeof parsedBody.lineChannelId === 'string'
    && typeof parsedBody.idempotencyKey === 'string'
    && parsedBody.event
    && typeof parsedBody.event === 'object'
    && !Array.isArray(parsedBody.event);
  if (!isEnvelope) return { present: false };

  var params = e && e.parameter ? e.parameter : {};
  var timestamp = String(params.bridgeTimestamp || '');
  var nonce = String(params.bridgeNonce || '');
  var signature = String(params.bridgeSignature || '');
  var secret = PropertiesService.getScriptProperties()
    .getProperty(CUSTOMER_LINE_BRIDGE_SECRET_PROPERTY) || '';
  if (secret.length < 32 || !/^\d{10}$/.test(timestamp)
      || !/^[A-Za-z0-9_-]{16,120}$/.test(nonce)
      || !/^[A-Za-z0-9+/=]{40,120}$/.test(signature)) {
    return { present: true, accepted: false, code: 'invalid_auth' };
  }
  var nowMillis = Date.now();
  if (Math.abs(nowMillis - Number(timestamp) * 1000) > 5 * 60 * 1000) {
    return { present: true, accepted: false, code: 'expired' };
  }
  var expected = Utilities.base64Encode(Utilities.computeHmacSha256Signature(
    timestamp + '.' + nonce + '.' + rawBody,
    secret,
    Utilities.Charset.UTF_8
  ));
  if (!_customerLineBridgeConstantTimeEqual(expected, signature)) {
    return { present: true, accepted: false, code: 'invalid_signature' };
  }
  if (!_customerLineBridgeReserveNonce(nonce, nowMillis)) {
    return { present: true, accepted: false, code: 'replay' };
  }
  if (parsedBody.event.type === 'bridge_test') {
    return { present: true, accepted: true, test: true, event: parsedBody.event };
  }
  var claim = _customerLineBridgeClaimEvent(parsedBody.idempotencyKey, nowMillis);
  if (!claim.accepted) {
    return { present: true, accepted: false, code: 'in_progress' };
  }
  return {
    present: true,
    accepted: true,
    test: false,
    duplicate: claim.duplicate,
    idempotencyKey: parsedBody.idempotencyKey,
    event: parsedBody.event
  };
}

function _customerLineDirectWebhookEnabled() {
  return PropertiesService.getScriptProperties()
    .getProperty(DIRECT_LINE_WEBHOOK_ENABLED_PROPERTY) !== 'false';
}

function configureCustomerLineBridge(secret) {
  if (typeof secret !== 'string' || secret.length < 32) {
    throw new Error('32文字以上のブリッジ秘密値が必要です。');
  }
  PropertiesService.getScriptProperties().setProperties({
    CUSTOMER_LINE_BRIDGE_SECRET: secret,
    LINE_DIRECT_WEBHOOK_ENABLED: 'true'
  });
  return { secretConfigured: true, directWebhookEnabled: true };
}

function disableDirectLineWebhookAfterCutover(confirmation) {
  if (confirmation !== 'DISABLE DIRECT LINE WEBHOOK AFTER CUTOVER') {
    throw new Error('確認文字列が一致しません。');
  }
  PropertiesService.getScriptProperties()
    .setProperty(DIRECT_LINE_WEBHOOK_ENABLED_PROPERTY, 'false');
  return { directWebhookEnabled: false };
}

function getCustomerLineBridgeStatus() {
  var props = PropertiesService.getScriptProperties();
  return {
    secretConfigured: (props.getProperty(CUSTOMER_LINE_BRIDGE_SECRET_PROPERTY) || '').length >= 32,
    directWebhookEnabled: _customerLineDirectWebhookEnabled()
  };
}

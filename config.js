const CONFIG = {
  // =============================
  // フォームの基本設定
  // =============================
  formTitle: "アフィリエイト申請フォーム",
  formDescription: "以下の手順に従って入力・作業を行ってください。",

  // =============================
  // アフィリエイトリンク設定
  // =============================
  affiliateUrl: "https://ここにアフィリエイトリンクを入力.com",
  affiliateButtonText: "アフィリエイトリンクを開く（必ずここから！）",

  // =============================
  // Google Apps Script の URL
  // GASを公開したらここに貼り付ける
  // =============================
  gasUrl: "https://script.google.com/macros/s/AKfycbzAxPvvIHvR-z8iYbZNDSIz4ZcJV4q3zFwTSC3SFsndeT0cmeVqiGa8VX9nqFRXhpYK/exec",

  // =============================
  // フォーム項目の設定
  // type: "text" / "select" / "textarea"
  // required: true / false
  // =============================
  fields: [
    {
      id: "name",
      label: "お名前",
      type: "text",
      required: true,
      placeholder: "例：山田太郎"
    },
    {
      id: "referrer",
      label: "紹介者名",
      type: "text",
      required: true,
      placeholder: "例：田中花子"
    }
    // 項目を追加する場合は上の {} ブロックをコピーして追加
    // ,{
    //   id: "phone",
    //   label: "電話番号",
    //   type: "text",
    //   required: false,
    //   placeholder: "例：090-1234-5678"
    // }
  ]
};

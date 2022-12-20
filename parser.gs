const getOgImage = (url) => {
  const response = UrlFetchApp.fetch(url);
  const data = response.getContentText();
  let ogImage;

  try {
    ogImage = Parser.data(data)
      .from('<meta property="og:image" content="')
      .to('" />')
      .build();
  } finally {
    if (!ogImage || !ogImage.match(/^http.*png$/)) {
      ogImage = "https://connpass.com/static/img/api/connpass_logo_1.png";
    }
  }

  return ogImage;
};

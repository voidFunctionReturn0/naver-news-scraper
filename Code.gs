function getNaverNews() {
  // 값 세팅
  const countOfFetchingNews = 100;
  const startRow = 2;
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const mainSheet = SpreadsheetApp.getActiveSheet();
  // let countOfLoad = 0;
  const titleColumnWidth = 300;
  const naverNewsSearchUrl = "https://openapi.naver.com/v1/search/news.json";
  const naverClientId = "Vgaw7chdxBKIzuu6LOmT";
  const naverClientSecret = "a2VDcbXszc";
  const apiOptions = {
    "headers": {
      "X-Naver-Client-Id": naverClientId,
      "X-Naver-Client-Secret": naverClientSecret,
    }
  }
  let result = [];

  // 스프레드시트에서 키워드 가져오기
  const lastKeywordRow = mainSheet.getLastRow();
  const keywords = mainSheet.getRange(startRow, 1, (lastKeywordRow-startRow+1), 1).getValues();
  Logger.log(keywords)

  // 키워드별로 처리
  for (let i = 0; i < keywords.length; i++) {
    keyword = keywords[i][0];

    // 새 뉴스 가져오기
    const newsSearchUrl = naverNewsSearchUrl + `?query=${keyword}&display=${countOfFetchingNews}`
    const responseJson = UrlFetchApp.fetch(newsSearchUrl, apiOptions);
    let newArticles = JSON.parse(responseJson.getContentText()).items;
    
    // 시트 세팅
    let keywordSheet = spreadSheet.getSheetByName(keyword);
    if (keywordSheet == null) {
      keywordSheet = spreadSheet.insertSheet();
      keywordSheet.setName(keyword);
      keywordSheet.getRange(1,1).setValue("발행일");
      keywordSheet.getRange(1,2).setValue("제목");
      keywordSheet.getRange(1,3).setValue("링크");
      keywordSheet.getRange(1,4).setValue("내용");
      keywordSheet.getRange("A1:D1").setBackground("#efefef");
      keywordSheet.setColumnWidth(2, titleColumnWidth);
    } else {
      newArticlesWithoutDeduplication = newArticles;
      newArticles = [];

      // 기존에 저장한 뉴스 가져오기
      const lastArticleRow = keywordSheet.getLastRow();
      let oldArticleLinks = keywordSheet.getRange(startRow, 3, (lastArticleRow-startRow+1), 1).getValues();

      // 기존 뉴스와 중복되지 않는 것만 새 뉴스로 인정함
      for (let j = 0; j < newArticlesWithoutDeduplication.length; j++) {
        const newArticle = newArticlesWithoutDeduplication[j];

        let isDuplicate = false;
        for (let k = 0; k < oldArticleLinks.length; k++) {
          const oldArticleLink = oldArticleLinks[k][0];

          if (oldArticleLink == newArticle.link) {
            isDuplicate = true
            break;
          }
        }
        if (isDuplicate == false) {
          newArticles.push(newArticle);
        }
      }
    }

    // 새 뉴스 저장
    if (newArticles != 0) {
      keywordSheet.insertRows(2, newArticles.length);
      for (let j = 0; j < newArticles.length; j++) {
        newArticle = newArticles[j];
        keywordSheet.getRange(j+2,1).setValue(new Date(newArticle.pubDate.substring(5,22)));
        keywordSheet.getRange(j+2,2).setValue(newArticle.title);
        keywordSheet.getRange(j+2,3).setValue(newArticle.link);
        keywordSheet.getRange(j+2,4).setValue(newArticle.description);
      }

      result.push(`${keyword} ${newArticles.length}개`);
    } 
  }
  // 실행 결과 기록
  spreadSheet.setActiveSheet(mainSheet);
  mainSheet.getRange(2,8).setValue(Utilities.formatDate(new Date(), "GMT+9", "yyyy. MM. dd. HH:ss"));
  mainSheet.getRange(3,8).setValue(`${result.join(", ")}`);
}
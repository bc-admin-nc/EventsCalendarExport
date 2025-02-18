interface DoGetEvent {
  parameter: { [key: string]: string };
  contextPath: string;
  contentLength: number;
  queryString: string;
}

function doGet(e: DoGetEvent): GoogleAppsScript.HTML.HtmlOutput {
  return HtmlService.createHtmlOutputFromFile("src/Index");
}
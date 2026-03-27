function triggerGitHubWorkflow() {
  const OWNER = 'John-Cena-DEV';
  const REPO = 'cdr-fetcher';
  const WORKFLOW_FILE = 'fetch_cdr.yml'; // exact filename
  const TOKEN = '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';

  const url = `https://api.github.com/repos/${OWNER}/${REPO}/actions/workflows/${WORKFLOW_FILE}/dispatches`;

  const payload = {
    ref: 'main'
  };

  const response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: `Bearer ${TOKEN}`,
      Accept: 'application/vnd.github+json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  Logger.log(response.getResponseCode());
}

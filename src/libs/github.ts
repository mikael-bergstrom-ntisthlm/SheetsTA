export namespace LibGithub {

  /**
   * Create a GitRepo object based on an URL
   * @param {string} url - the Github URL to be interpreted
   * @returns {GitRepo | undefined} a GitRepo object if URL was valid; otherwise returns undefined
   */
  export function InterpretURL(url: string):
    GitRepo | undefined {

    const re = new RegExp("https?:\/\/.*github.com\/(?<user>[^/]+)\/(?<repo>[^/]+)\/*.*$");

    let result = re.exec(url)?.groups;
    if (!result) return undefined;

    return {
      user: result['user'],
      name: result['repo']
    };
  }

  /**
   * Sanitize a GitHub URL
   * @param {string} origUrl - The original "dirty" URL
   * @returns {string} The resulting "clean" URL
   */
  export function UrlSanitize(origUrl: string): string {
    let repo = InterpretURL(origUrl);

    return repo == undefined ? origUrl : BuildWebURL(repo);
  }

  /**
   * Get an array of when commits were made to a GitHub repo
   * @param {GitRepo} repo - The repository to examine
   * @param {string} userEmail - Filter by this E-mail, if defined
   * @returns {Date[]} An array of dates where commits were made
   */
  export function GetCommitDates(repo: GitRepo, userEmail?: string): Date[] {
    let url = BuildApiRepoURL(repo) + "/commits";

    let editTimestamps: Date[] = [];

    let response = UrlFetchApp.fetch(url);
    if (response.getResponseCode() == 200) {
      let commits: Commit[] = JSON.parse(response.getContentText()) as Commit[];
      if (!commits) return [];

      commits.forEach(commit => {
        if (userEmail === undefined || commit.commit.author.email === userEmail) {
          editTimestamps.push(new Date(commit.commit.author.date));
        }
      });

      return editTimestamps;
    }
    return [];
  }

  function BuildApiRepoURL(repo: GitRepo) {
    return `https://api.github.com/repos/${repo.user}/${repo.name}`;
  }

  export function BuildWebURL(repo: GitRepo) {
    return `https://github.com/${repo.user}/${repo.name}`;
  }

  interface GitRepo {
    user: string
    name: string
  }

  interface Commit {
    commit: {
      message: string
      author: {
        email: string
        name: string
        date: string
      }
      committer: {
        email: string
        name: string
        date: string
      }
    }
  }
}
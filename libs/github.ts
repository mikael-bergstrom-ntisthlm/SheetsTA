namespace GithubTA {

  export function InterpretURL(url: string): GitRepo | undefined {

    const re = new RegExp("https?:\/\/.*github.com\/(?<user>[^/]+)\/(?<repo>[^/]+)\/*.*$");
    // https?:\/\/.*github.com\/(?<user>.+?)\/(?<repo>.+?)[\/$]*

    let result = re.exec(url)?.groups;
    if (!result) return undefined;

    return {
      user: result['user'],
      name: result['repo']
    };
  }

  export function UrlSanitize(origUrl: string): string {

    let repo = InterpretURL(origUrl);
    
    return repo == undefined ? origUrl : BuildWebURL(repo);
  }

  export function GetCommitDates(repo: GitRepo, userEmail?:string): Date[] {
    let url = BuildApiRepoURL(repo) + "/commits";

    let editTimestamps: Date[] = [];

    let response = UrlFetchApp.fetch(url);
    if (response.getResponseCode() == 200) {
      let commits: Commit[] = JSON.parse(response.getContentText()) as Commit[];
      if (!commits) return [];

      commits.forEach(commit => {
        if (userEmail === undefined || commit.commit.author.email === userEmail)
        {
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
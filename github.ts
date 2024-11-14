namespace GithubTA {
  export function UrlSanitize(origUrl: string): string {
    // origUrl = "https://github.com/TE22A-John-Julius/JuliusDiagnos/blob/main/Diagnos_Julius_John_TE22-s/Program.cs";

    const re = new RegExp("https?:\/\/github.com\/(?<user>.+?)\/(?<repo>.+?)[/.].*");

    let result = re.exec(origUrl)?.groups;
    if (!result) return origUrl;

    return URLBuilder(result['user'], result['repo']);
  }

  let URLBuilder = (user: string, repo: string) =>
    `https://www.github.com/${user}/${repo}`;
}
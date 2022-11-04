// export class GithubConnect {
//   constructor() {

//   }

//   githubUsersUrl = 'https://api.github.com/users/'
//   userDataKeys = ["login", "name", "location", "bio"];

//   async function fetchUserData(userName) {
//     // fetch the user's data from the GitHub API.
//     const url = `${githubUsersUrl}${userName}`;
//     const obj = await fetchFrom(url, `github user ${userName} not found`);
//     // prepare the data for the table.
//     const userData = userDataKeys.map((key) => obj[key] || "");
//     console.log(`fetchUserData`, userData);
//     return userData;
//   }
// }

const githubUsersUrl = "https://api.github.com/users/";
const userDataKeys = ["login", "name", "location", "bio"];

export async function fetchUserData(userName) {
  // fetch the user's data from the GitHub API.
  const url = `${githubUsersUrl}${userName}`;
  const obj = await fetchFrom(url, `github user ${userName} not found`);
  // prepare the data for the table.
  const userData = userDataKeys.map((key) => obj[key] || "");
  console.log(`fetchUserData`, userData);
  return userData;
}

/**
 * Fetch data from a URL
 * @param {*} url
 * @returns promise that resolves to the JSON object returned by the url
 */
 async function fetchFrom(url, errorMessage) {
  const response = await fetch(url);
  if (!response.ok) {
    //console.log(response);
    throw new Error(`${errorMessage}`);
  }
  const data = await response.json();
  return data;
}
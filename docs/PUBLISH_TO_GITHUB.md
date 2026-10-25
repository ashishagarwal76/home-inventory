# Post this project on GitHub

1. Extract the downloaded package. Open the extracted project folder so you can see `README.md`, `LICENSE`, `apps-script/`, `scanner/`, and the other included folders.
2. Sign in to GitHub and open [New repository](https://github.com/new).
3. Choose your account as the owner. A suggested repository name is `home-inventory`.
4. Add a description, such as `Google Sheets home inventory with photos, nested boxes, and QR labels.` Choose **Public** if you want anyone to be able to download the reusable code.
5. Leave the options for generating a README, license, and .gitignore unselected; those files are already in this package. Click **Create repository**. [GitHub's repository creation instructions](https://docs.github.com/en/repositories/creating-and-managing-repositories/creating-a-new-repository) describe these controls.
6. On the empty repository page, use the **uploading an existing file** link. In a repository with files, use **Add file → Upload files**.
7. Drag the extracted project **contents** into the upload area, preserving their folders. `README.md` should be at the repository root. Uploading the source this way makes it browseable on GitHub.
8. Use a commit message such as `Add reusable home inventory app`, then save the upload with **Commit changes** or the commit/proposal control shown by GitHub. [GitHub's upload guide](https://docs.github.com/en/repositories/working-with-files/managing-files/adding-a-file-to-a-repository) covers the upload screen.
9. Open the repository's README and follow its Setup link to check the instructions. Others can download the source through **Code → Download ZIP** and connect their own Google files.

The project contains fewer than 100 files, and every file is smaller than the web uploader's per-file limit. Keep the root dotfiles `.gitignore` and `.gitattributes` when uploading.

This publishes the reusable source and empty template. Each person's working inventory app is deployed through their Google Apps Script project, following [Setup](SETUP.md). Live installation, camera, and print checks are listed in [Validation](VALIDATION.md).

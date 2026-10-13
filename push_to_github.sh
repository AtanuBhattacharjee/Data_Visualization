#!/usr/bin/env bash
# Update the existing GitHub repository with the new version of the app.
# Run from inside this folder after replacing YOUR-GITHUB-USERNAME:
#       bash push_to_github.sh
# (For a brand-new repository, create it empty on github.com first.)

GH_USER="YOUR-GITHUB-USERNAME"
REPO="session02-frequency-app"

git clone "https://github.com/${GH_USER}/${REPO}.git" repo_tmp
cp app.R manifest.json hypertension-dataset.xlsx README.md .gitignore repo_tmp/
cd repo_tmp
git add app.R manifest.json hypertension-dataset.xlsx README.md .gitignore
git commit -m "Add the slide builder: recommended displays for uploaded data, PowerPoint and PDF export"
git push origin main
cd .. && rm -rf repo_tmp

@echo off
git init
git add .
git commit -m "Initial commit"
git remote remove origin
git remote add origin https://github.com/iamholm/AppObzorko2.git
git push -u origin master 
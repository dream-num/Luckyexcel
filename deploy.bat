# deploy Demo
gulp build
cd dist
git init
git remote add origin https://github.com/mengshukeji/LuckyexcelDemo.git
git add .
git commit -m 'deploy Luckyexcel demo'
git push -f origin master:gh-pages



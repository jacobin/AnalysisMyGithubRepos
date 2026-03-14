#!/bin/bash
# 把github上jacobin所有的项目信息都导出到以tab分割的csv文件里
gh repo list jacobin -L10000 > /d/AllMyGithubRepos.txt
# gh repo list jacobin --json updatedAt,url --jq .[].updatedAt,url -L1000

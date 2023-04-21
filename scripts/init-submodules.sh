#!/bin/bash

cd "$(dirname "${BASH_SOURCE[0]}")"/..

git config -f .gitmodules --get-regexp '^submodule\..*\.path$' |
  while read path_key path
  do
    name=$(echo $path_key | sed 's/\submodule\.\(.*\)\.path/\1/')
    url_key=$(echo $path_key | sed 's/\.path/.url/')
    branch_key=$(echo $path_key | sed 's/\.path/.branch/')
    url=$(git config -f .gitmodules --get "$url_key")
    branch=$(git config -f .gitmodules --get "$branch_key" || echo "main")

    git submodule add -b $branch --name $name $url $path || continue
  done

if $(git rev-parse --is-inside-work-tree &> /dev/null); then
  git submodule update --init --recursive
  git submodule foreach 'git checkout main'
  git submodule foreach 'git pull'
  git submodule update --remote --merge
fi

#!/bin/bash

while read entry ; do
  argarray=($entry)
  username=${argarray[0]}
  count=${argarray[1]}
  echo $username
  echo $count
  ruby ../harvest.rb ${username} ${count}
done <list
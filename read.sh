#!/bin/bash
# ls -d $PWD/*.docx >1.txt
# ls -R | grep .docx$ >1.txt
find  $PWD | xargs ls -d | grep .docx$ >1.txt
python3 read-docx.py >2.txt
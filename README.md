Grep keyword in Microsoft Powerpoint (pptx) files
===

[![MIT License](http://img.shields.io/badge/license-MIT-blue.svg?style=flat)](LICENSE)

## Summary

Find keywords in pptx files. It will search pptx files recursively from the current working directory.

## Build

    $ dmd pptxgrep.d
    $ alias pptxgrep=$PWD/pptxgrep # or put them somewhere in $PATH

## Usage

    $ pptxgrep keyword
    Found "keyword" in hoge/hoge.pptx at slide 4
    Found "keyword" in hoge/hoge.pptx at slide 1
    Found "keyword" in test.pptx at slide 3

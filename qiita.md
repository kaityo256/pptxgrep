# パワーポイント内のテキストをgrepする(D言語版)

## はじめに

大量のパワーポイント資産がある時、「あれ？この話どこでしたっけ？」とか「あの話をしたスライドどれだっけ？」と悩むことが多い。そんな時のために[pptxファイル内をgrepするRubyスクリプト](https://qiita.com/kaityo256/items/2977d53e70bbffd4d601)を書いた。これはこれで便利だったのだが、手抜き処理が多かったために動作が遅く、大量のpptxファイルの検索には向かなかった。Rubyスクリプトのまま高速化することもできるのだろうが、勉強を兼ねて別の言語で書き直したい。というわけでD言語版を作った。

ファイルは
https://github.com/kaityo256/pptxgrep
においてある。

## 使い方

とりあえずコンパイルして、aliasなりPATHの通ったところに置くなりする。

```shell
$ dmd pptxgrep.d
$ alias pptxgrep=$PWD/pptxgrep # or put them somewhere in $PATH
```

実行する。引数は検索したいキーワードのみ。実行したディレクトリから再帰的にサブディレクトリを探しに行き、見つけた*.pptxファイルの中で、キーワードを含むスライド番号を出力する。

```shell
$ pptxgrep keyword
Found "keyword" in hoge/hoge.pptx at slide 4
Found "keyword" in hoge/hoge.pptx at slide 1
Found "keyword" in test.pptx at slide 3
```

なお、ファイルのパスについてはソートされているが、出力されるスライド番号についてはソートされないため、番号が前後することがある。

## 動作原理

PPTXファイルはzip圧縮されたXMLファイルなので、unzipして出てきたXMLファイル内を検索すればよろしい。というわけで以下のような処理ができれば良い。

1. カレントディレクトリ以下の*.pptxを再帰的に検索する
1. pptxファイルを見つけたらその中身を調べる
1. `ppt/slides/slide([0-9]+).xml`にマッチするファイルがあればそれをメモリに展開する
1. 展開したデータを文字列として解釈し、検索ワードを含んでいたら、ファイル名とスライド番号を出力する


### カレントディレクトリからファイルを再帰的に検索する

`std.file.dirEntries`という、そのものずばりの関数があるのでそれを使う。

```d
auto cwd = std.file.getcwd();
auto d = dirEntries(cwd,"*.pptx",SpanMode.depth);
```

カレントディレクトリを`std.file.getcwd()`で取得し、それを`std.file.dirEntries`にわたす。検索方法は`SpanMode`で指定するが、あとでソートするのでなんでも良い。

これで得られるファイルはソートされていないので、`string []`に変換してソートする。

```d
  string [] files;
  foreach(string filename; d){
    files ~= filename;
  }
  files.sort!();
```

`map`とかで一発で書けるんだろうけど気にしない。これでサブディレクトリ以下にある全ての*.pptxファイルを文字列配列として得ることができた。

### pptxファイルの中身を調べる

pptxはzipファイルなので、その内容物を調べるのに`std.zip.ZipArchive`を使う。

```d
auto zip = new ZipArchive(read(filename));
```

zipファイルの中の`ppt/slides/slide([0-9]+).xml`にマッチするファイルがスライドなので、それを探す。

```d
foreach (name, am; zip.directory)
  {
    foreach(m; match(name, r"ppt/slides/slide([0-9]+).xml$"))
    {
     // スライドファイルを見つけた
    } 
  }
```

ちなみに、カッコで囲んだ部分がスライド番号なので、あとで使う。

### スライドファイルを展開し、文字列に変換する

ZIPファイル名のエントリーは`ArchiveMember`として受け取れる。これを個別にメモリ上でunzipするには、`ZipArchive.expand`を使えば良い。

```d
      zip.expand(am);
```

すると、`expandedData`メンバに展開される。これは`ubyte []`なので、これを文字列に変換する。

```d
char *cstr = cast(char*)am.expandedData;
auto len = am.expandedData.length;
string str = cast(string) cstr[0..len];
```

これで`str`にスライド一枚のXMLファイルがテキスト形式で格納された。

### スライドに検索ワードが含まれていたらファイル名とスライド番号を出力

```d
if(match(str, keyword)){
    auto rname = relativePath(filename);
    writefln("Found \"%s\" in %s at slide %s",keyword,rname, slidenum);
}
```

難しいところは無いと思う。ただ、ファイル名が絶対パスになっているのが不便だったので、`relativePath`でカレントディレクトリからの相対パスを出力するようにしている。

### 速度

自分のすべてのスライド資産にたいして、キーワード「hoge」を検索するのにかかった時間をRuby版と比較してみる。

```
$ time ruby ~/github/grep_pptx/grep_pptx.rb hoge > /dev/null
ruby ~/github/grep_pptx/grep_pptx.rb hoge > /dev/null  102.04s user 12.08s system 97% cpu 1:57.60 total

$ time ~/github/pptxgrep/pptxgrep hoge > /dev/null
~/github/pptxgrep/pptxgrep hoge > /dev/null  0.24s user 0.40s system 99% cpu 0.645 total
```

Ruby版が2分弱、D言語版は0.65秒ということで、100倍以上早くなりましたな。
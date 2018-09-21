# パワーポイント内のテキストをgrepする(D言語版)

## はじめに

大量のパワーポイント資産がある時、「あれ？この話どこでしたっけ？」とか「あの話をしたスライドどれだっけ？」と悩むことが多い。そんな時のために[pptxファイル内をgrepするRubyスクリプト](https://qiita.com/kaityo256/items/2977d53e70bbffd4d601)を書いた。これはこれで便利だったのだが、手抜き処理が多かったために動作が遅く、大量のpptxファイルの検索には向かなかった。Rubyスクリプトのまま高速化することもできるのだろうが、勉強を兼ねて別の言語で書き直したい。というわけでD言語版を作った。

ファイルは
https://github.com/kaityo256/pptxgrep
においてある。

自分で言うのもなんだけどわりと便利なので、pptx資産が大量にある人は是非使って欲しい。

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

### XMLからテキストを抽出

XMLファイルがテキスト形式で`str`に格納されたので、あとは`match`なり`indexOf`なりでキーワードを含むか調べたくなるが、[前に書いた](https://qiita.com/kaityo256/items/2977d53e70bbffd4d601#%E3%83%86%E3%82%AD%E3%82%B9%E3%83%88%E3%81%AE%E6%8E%A2%E3%81%97%E6%96%B9)とおり、スライド上では一続きの文字列に見えても、XMLではバラバラに格納されていることがある。例えば「平成30年」という言葉は、「平成」「30」「年」でバラバラになっており、このままだと「平成30年」でヒットしない。そのためにXMLのテキストノードをすべて抽出し、結合した文字列に対して検索をかける必要がある。

素直にD言語標準のXMLのパーサー`std.xml`を使うとこんな感じに書けるだろう。

```d
auto xml = new DocumentParser(str);
dstring text;
xml.onText = (string s)
{
  text ~= s.to!dstring;
};
xml.parse();
```

これは`DocumentParser`に「テキストノードを見つけた時」のイベントハンドラを登録しておいて、`xml.parse()`を呼び出すと、後は見つけるたびにテキストが追加されていく仕組み。こうして得られた`text`に対して`indexOf`をかければ良いのだが、残念ながらこのコードは遅い。おそらく`std.xml`のパーサーが遅いのだと思われる。仕方ないので、テキストノードを抽出するコードを自分で書こう。

テキストは`<a:t>`と`</a:t>`に囲まれているので、それを抽出すれば良い。素直に書けばこんな感じになるだろうか。

```d
dstring extractText(string xmltext)
{
  dstring dxml = xmltext.to!dstring();
  dstring text;
  while(findSkip(dxml, "<a:t>")){
    auto e = indexOf(dxml,"</a:t>");
    text ~= dxml[0..e];
  }
  return text;
}
```

あとはこいつを

```d
dstring text = extractText(str);
```

と呼び出せば、`text`にテキストノードが全て結合されたものが入る。

### スライドに検索ワードが含まれていたらファイル名とスライド番号を出力

```d
if(text.indexOf(keyword) !=-1)
{
  auto rname = relativePath(filename);
  writefln("Found \"%s\" in %s at slide %s",keyword,rname, slidenum);
}
```

そのままなので難しいところは無いと思う。ただ、ファイル名が絶対パスになっているのが不便だったので、`relativePath`でカレントディレクトリからの相対パスを出力するようにしている。

### 速度

自分のすべてのスライド資産にたいして、キーワード「hoge」を検索するのにかかった時間をRuby版と比較してみる。何度か実行して、ファイル情報がキャッシュにのった状態で測定。

```
$ time ruby ~/github/grep_pptx/grep_pptx.rb hoge > /dev/null
ruby ~/github/grep_pptx/grep_pptx.rb hoge > /dev/null  80.60s user 6.74s system 96% cpu 1:30.16 total

$ time ~/github/pptxgrep/pptxgrep hoge > /dev/null
~/github/pptxgrep/pptxgrep hoge > /dev/null  2.13s user 0.62s system 92% cpu 2.974 total
```

Ruby版が90秒、D言語版は3秒ということで、30倍くらい早くなった。なお、Ruby版は

* unzipをシェルで呼び出している
* 必要ないファイルもすべて展開し、ファイルに吐いている
* XPathを使ったXML解析をしている

というハンデがあるので、これはフェアな比較になっていないことに注意。

ちなみにD言語版でも、XMLでパースした場合は

```
$ time ~/github/pptxgrep/pptxgrep hoge > /dev/null
~/github/pptxgrep/pptxgrep hoge > /dev/null  19.15s user 0.85s system 95% cpu 20.937 total
```

と、自前パース版に比べて7倍くらい遅くなる。

## まとめ

D言語でpptx内テキストのgrepコマンドを作った。数十行足らずでこれだけのことができるんだから大したもんだと思う。ただし、`std.xml`が遅いのはちょっと困る。これは開発コミュニティも把握しているっぽいが、`std.xml2`は[Abandoned](https://wiki.dlang.org/Review_Queue)になってますね・・・
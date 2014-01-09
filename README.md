# VBAFramework

VBAFramework is a bunch of the .NET Framework Clone Classes and more for .NET programmers.  
VBAFrameworkは、VBAっぽい書き方なんてわざわざ覚えたくないという.NETプログラマ(作者)のための  
.NET Frameworkクローンライクなクラス群です。

## Features
* 出来る限り.NET Frameworkと同じ書き方ができるようにしてます
* 参照設定を一切不要にしています
* 依存関係を一切持たず、各クラスファイル単体で動作します
* 32bit/64bit Office 両対応（のはず）です

## What is good
* クラスファイルなので、VBEでメソッドを入力する際、入力補完が効きます
* 必要なクラスファイルだけインポートすれば即使えます。らくちんです(自分評)。

## What is bad
* 本家のメソッドを網羅してない (作者が実務で必要になった分だけ実装...)
* エラー処理が雑（というかしてない...）

## Contents
- [.NET Frameworkクローンライクなクラス群] VBA_Clipboard.cls, VBA_Directory.cls, VBA_File.cls, VBA_Path.cls, VBA_Stopwatch.cls, VBA_StringBuilder.cls
- [オレ得クラス] VBA_ConfigFile.cls (簡易Configファイル作成クラス), VBA_VAMIE.cls (IEの自動制御用クラス)
- [サンプルコード] Sample.bas, VBA_Static.bas

## Usage
Please see the [Sample.bas](https://github.com/dck-jp/VBAFramework/blob/master/Sample.bas).  
使い方は、[Sample.bas](https://github.com/dck-jp/VBAFramework/blob/master/Sample.bas)を見てください。

.NET Frameworkが使えるのであれば、だいたいフィーリングで大丈夫だと思います。  
.NET Frameworkを知らない場合は、がんばってF#を勉強すると良いことあるかも？

### Tips
.NET Frameworkで Staticなクラスに対応するVBAFrameworkのクラスに関しては、  
適当なモジュール(例えば、[VBA_Statics.bas](https://github.com/dck-jp/VBAFramework/blob/master/VBA_Static.bas))内で  
Publicな変数を作成して宣言と初期化を同時にやると良いんじゃないでしょうか。

当該変数を介して、いつでも・どこからでも参照できるようになり、  
あたかもStaticクラスのように扱えるようになるので個人的にはオススメです。


## License
This source code is under [MIT License](https://github.com/dck-jp/VBAFramework/blob/master/LICENSE)  
ソースコードは[MIT License](https://github.com/dck-jp/VBAFramework/blob/master/LICENSE)で配布しています。

てきとうによろしくどうぞー。

## Author
D*isuke YAMAKAWA ([@d_ymkw](https://twitter.com/d_ymkw))

PDF Exporter for Office
========

Microsoft Office(Word Excel PowerPoint)で簡単にPDFを作成するアドインです。  
パスワード付きPDFも作成できます。  

ダウンロードは[releases](https://github.com/t-miyake/PdfExporterForOffice/releases)からできます。

オープンソースかつ無料でご利用いただけますが、無サポート、無保証です。  
専用のカスタマイズやサポートが必要な場合は、個別にご相談ください。  

アドイン(リボンにアイコンが追加されます)  
![Screenshot](https://github.com/t-miyake/PdfExporterForOffice/blob/master/Screenshots/Screenshot.png) 


## 使い方 (Word Excel PowerPoint 全て共通)  
簡単にするために、細かい設定などは全て省いています。  
(設定内容は次項に列記)

### ・PDF(パスワードなし)の作成  
1. リボンの「アドイン」内の「PDF Exporter for xxx」にある「Export PDF」ボタン を押します。
1. ファイルの保存先の確認画面が表示されるので、好きな名前を付けて保存します。
1. 作成完了(∩´∀｀)∩   
 
### ・PDF(パスワードあり)の作成  
1. リボンの「アドイン」内の「PDF Exporter for xxx」にある「Export PDF (Password)」ボタン を押します。
1. パスワードの入力画面が表示されるので、パスワードを入力します。(2回同じものを入力)
1. ファイルの保存先の確認画面が表示されるので、好きな名前を付けて保存します。
1. 作成完了(∩´∀｀)∩  
  
## PDFの出力設定
以下の設定に固定されています。(変更できません)

### ・PDFのパスワード設定 (パスワード付きのPDFを作成する場合)
1. 文書を開くパスワード(ユーザパスワード)と権限パスワード(マスターパスワード)は、どちらも同じものを設定。
1. テキスト、画像、およびその他の内容のコピーは有効(許可)。
1. 印刷は有効(許可)。

### ・WordのPDF出力設定  
1. PDFの品質は、印刷に最適化。  
1. PDF化の対象範囲は、全てのページ。  
1. 変更履歴とコメントは、全て含まない。  
1. ブックマークは、作成しない。  
1. ドキュメントのプロパティは、含める。  
1. アクセシビリティ用のドキュメント構造タグは、含める。  
1. PDF/A 準拠は、しない。  
1. フォントの埋め込みが不可能な場合はテキストをビットマップに変換する。  

### ・ExcelのPDF出力設定
1. PDFの品質は、印刷に最適化。
1. PDF化の対象範囲は、開いているシート。
1. ドキュメントのプロパティは、含める。
1. 印刷範囲は、無視しない。(印刷設定に従ってPDF化)

### ・PowerPointのPDF出力設定
1. PDFの品質は、印刷に最適化。
1. PDF化の対象範囲は、全てのスライド。
1. PDF化の対象は、スライド。(配布資料等ではない)
1. スライドに枠は、付けない。
1. 非表示のスライドは、含めない。
1. コメント及びインク注釈は、全て含めない。
1. ドキュメントのプロパティは、含める。
1. アクセシビリティ用のドキュメント構造タグは、含める。
1. PDF/A 準拠は、しない。
1. フォントの埋め込みが不可能な場合はテキストをビットマップに変換する。


## 利用ライブラリ
以下のライブラリを利用しています。

* iTextSharp
    - [GNU Affero General Public License](http://www.gnu.org/licenses/agpl.html)
    - https://www.nuget.org/packages/iTextSharp/

## 既知の不具合
 1. パスワード付きのPDFを作成すると、アプリケーション(Word等)の終了が少し遅くなる。  
 1. Excelで何もないシートでパスワード付きのPDFを作成すると、開けないPDFが作成される。

## ライブラリを読み込むには
```vbscript
thisPath = left(wscript.scriptfullname, len(wscript.scriptfullname) - len(wscript.scriptname))
Execute(CreateObject("Scripting.FileSystemObject").OpenTextFile(thisPath + "VBS\__init__.vbs").ReadAll())
```
この 2 行は、他のプログラミング言語でいう「指定include」や「import」に相当します。
記述するだけで、ライブラリ内のすべての関数がスクリプト内で使用できるようになります。
各関数の詳しい使い方は、それぞれの .vbs ファイルのコメントを参照してください。

このライブラリは自由に使用・改変・拡張していただいて構いません。

## ライブラリの構成

このライブラリは以下の設計原則に従っています

【1関数1ファイル】
  各 .vbs ファイルは 1 つの関数のみを定義します。
  ファイル名と関数名は常に一致します。

【public（公開） と private（非公開）の分離】
  * public  : char_code/, directory/, excel/ ディレクトリ内
             使い手が直接呼び出すため、詳細なドキュメント（英語＋日本語）を含みます。
  * private : char_code/private/, directory/private/ ディレクトリ内
             内部実装向けで、使い手には参照されません。
             コメントは一切なく、高いコード読解能力またはAI支援を前提とします。

【メンテナンス】
  private 領域のコードは、複雑なロジックを効率的に実装しています。
  修正や拡張には高度なコード読解スキルが必要です。
  生成AI（例：GitHub Copilot）の支援を受けることを推奨します。


## How to load the library

Add the following two lines to the top of your script
```vbscript
thisPath = left(wscript.scriptfullname, len(wscript.scriptfullname) - len(wscript.scriptname))
Execute(CreateObject("Scripting.FileSystemObject").OpenTextFile(thisPath + "VBS\__init__.vbs").ReadAll())
```
This works like "include" or "import" in other programming languages.
Once these two lines are in place, all library functions become available in your script.
The documentation for each function is written in comments inside each .vbs file.

Feel free to use, modify, or extend this library however you like.

## Library Structure

This library follows these design principles:

【One Function Per File】
  Each .vbs file defines exactly one function.
  The filename and function name always match.

【Public and Private Separation】
  * public  : Files in char_code/, directory/, excel/ directories
             Called directly by users, includes detailed documentation (English + Japanese).
  * private : Files in char_code/private/, directory/private/ directories
             For internal implementation; not referenced by users.
             No comments $2014 assumes high code comprehension ability or AI assistance.

【Maintenance】
  Code in the private areas is complex and optimized for efficiency.
  Modifications or extensions require advanced code reading skills.
  AI assistance (e.g., GitHub Copilot) is recommended.

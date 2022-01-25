Rem https://github.com/cpp-stdio/vbs_plugins.git からデータをコミットします。

@echo on

Set FILE_NAME=vbs_plugins
Set BRANCH=master
Set GIT_HTTP_PROXY=github.com/cpp-stdio/vbs_plugins.git
Set GIT_PASSWORD=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

cd %~dp0%
call %~dp0%GetCurrentTime\Pattern1.bat

@echo %CURRENT_TIME%に開始します

cd ../../

If Exist %FILE_NAME% (
    Rem 環境変数の遅延展開
    setlocal enabledelayedexpansion

    Set MESSAGE=%FILE_NAME% was committed ^in %CURRENT_TIME%

    cd %FILE_NAME%
    git add .
    Rem git status

    git commit -m "!MESSAGE!"
    git push origin %BRANCH%:%BRANCH%
) Else (
    echo %FILE_NAME%というフォルダがないため、コミットできませんでした
)

Rem 全自動という用途ではないため
PAUSE

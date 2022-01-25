Rem https://github.com/cpp-stdio/English.git からデータをクローンします。

@echo on

Set FILE_NAME=English
Set BRANCH=main
Set GIT_HTTP_PROXY=github.com/cpp-stdio/English.git
Set GIT_PASSWORD=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

cd %~dp0%
cd ../../

If Exist %FILE_NAME% (
    cd %FILE_NAME%
    git pull %BRANCH%
    echo %FILE_NAME%を更新しました。
) Else (
    git clone https://cpp-stdio:%GIT_PASSWORD%@%GIT_HTTP_PROXY% %FILE_NAME%
    cd %FILE_NAME%
    git checkout -b %BRANCH%
    echo %FILE_NAME%をクローンしました。
)

Rem 全自動という用途ではないため
PAUSE

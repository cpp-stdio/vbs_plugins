Rem https://github.com/cpp-stdio/vbs_plugins.git ����f�[�^���N���[�����܂��B

@echo on

Set FILE_NAME=vbs_plugins
Set BRANCH=master
Set GIT_HTTP_PROXY=github.com/cpp-stdio/vbs_plugins.git
Set GIT_PASSWORD=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

cd %~dp0%
cd ../../

If Exist %FILE_NAME% (
    cd %FILE_NAME%
    git pull %BRANCH%
    echo %FILE_NAME%���X�V���܂����B
) Else (
    git clone https://cpp-stdio:%GIT_PASSWORD%@%GIT_HTTP_PROXY% %FILE_NAME%
    cd %FILE_NAME%
    git checkout -b %BRANCH%
    echo %FILE_NAME%���N���[�����܂����B
)

Rem �S�����Ƃ����p�r�ł͂Ȃ�����
PAUSE

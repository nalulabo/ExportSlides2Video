@echo off
where /Q pandoc
if %errorlevel% == 0 (
    pandoc -o custom-reference.pptx --print-default-data-file reference.pptx
) else (
    @echo on
    echo ���s�ɂ�Pandoc���K�v�ł��B
    echo https://pandoc.org/installing.html
    pause
    exit
)

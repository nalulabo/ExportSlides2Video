@echo off
where /Q pandoc
if %errorlevel% == 0 (
    pandoc slides.md -o presentation.pptx --reference-doc=custom-reference.pptx
    if %errorlevel% == 0 (
        cscript ../ExportSlides2Video.vbs presentation.pptx
        goto EOF
    ) else (
        echo "not green."
        goto EOF
    )
) else (
    @echo on
    echo 実行にはPandocが必要です。
    echo https://pandoc.org/installing.html
    goto EOF
)

EOF:
    pause
    exit

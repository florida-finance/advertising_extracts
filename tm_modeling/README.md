# Instructions for compiling

1) conda create --name myenv python
2) activate myenv (on windows this is "source activate myenv")
3) pip install -r requiremnets.txt
4) pyi-makespec --onefile [FILE NAME]
5) Add to spec: 
>import sys <br>
>sys.setrecursionlimit(5000)
6) pyinstaller [FILE NAME].spec


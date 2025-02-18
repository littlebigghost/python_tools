# python_tools

## tool list
* url2word

## how to use

### Use source code
1. install dependencies
```shell
pip install -r requirements.txt
```
2. run
```shell
python url2word.py
```

### Use exe

download [v1.0 exe](https://github.com/littlebigghost/python_tools/releases/tag/v1.0)

## package exe

1. install pyinstaller
```shell
pip install pyinstaller
```
2. package exe
```shell
pyinstaller --onefile --add-data "D:\xxx\Lib\site-packages\newspaper\resources;newspaper/resources" --noconsole main.py
```
--add-data: package dependencies
--noconsole: don't show console

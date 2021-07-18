# Fxcel - Excel operations library  

![Fxcel](https://raw.githubusercontent.com/tatsuya-midorikawa/Fxcel/main/assets/fxcel.png)  


## What's this?  

- Fxcel は F# で簡単に Excel の COM 操作をするためのライブラリです。  
  - C# 向けの Excel COM 操作ライブラリである ***[Midoliy.Office.Interop.Excel](https://github.com/Midoliy/Midoliy.Office.Interop.Excel)*** のラッパーライブラリとなります。
- .NET 5.0 以上の環境をサポートしています。  
- 主に F# Script や F# Interactive での利用を想定して設計をしていますが、Console アプリや Desktop アプリでも問題なく利用可能です。  
- COM を利用するため Excel のインストールが必要です。  


## Get started  

### 1. F# Scriptで利用する

#### 1-1. **.fsx** ファイルを作成する  

まずはコーディングを始めるために **main.fsx** を作成して、VSCode で開きましょう。  

```powershell
mkdir D:/work
cd D:/work
new-item main.fsx
code D:/work
```

#### 1-2. Fxcel を読み込む

**main.fsx** に Fxcel を利用するためのコードを追加します。

```fsharp
#r "nuget: Fxcel"
open Fxcel
```  

### 2. F# プロジェクトで利用する

#### 2-1. 新規プロジェクトを作成する  

```powershell
mkdir D:/work
cd D:/work
dotnet new console -lang=F# -o=FxcelSample
``` 
#### 2-2. Fxcel を読み込む 

```powershell
cd D:/work/FxcelSample
dotnet add package Fxcel
``` 


## Reference
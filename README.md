# Pin-WindowPosition

ウインドウタイトルを指定する場合

```
Pin-WindowPosition <ウインドウタイトル>
```

ウインドウタイトルとexe名を指定する場合
```
Pin-WindowPosition <ウインドウタイトル> <exe名>
```
### 制限事項

- ウインドウタイトル、またはウインドウタイトルとexeの名称が複数同時に起動するケースは考慮していません。
- 終了するためには `taskkill /im Pin-WindowPosition.exe /f` を実行してください

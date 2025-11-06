# Windows Server 2022 における OpenSSH SCP 専用ユーザー設定手順

## 1. OpenSSH Server のインストール確認

Windows Server 2022 では OpenSSH Server
は既定でインストールされていないため、利用するにはインストールが必要。

### インストール

``` powershell
Add-WindowsCapability -Online -Name OpenSSH.Server~~~~0.0.1.0
```

### インストール確認

``` powershell
Get-WindowsCapability -Online | Where-Object Name -like 'OpenSSH.Server*'
```

出力例:

    Name  : OpenSSH.Server~~~~0.0.1.0
    State : Installed

------------------------------------------------------------------------

## 2. サービス起動/自動起動設定/待ち受け確認

``` powershell
Start-Service sshd
Set-Service -Name sshd -StartupType Automatic
netstat -ano | findstr :22
```

出力例:

    TCP    0.0.0.0:22    0.0.0.0:0    LISTENING    1072
    TCP    [::]:22       [::]:0       LISTENING    1072

※1072 はプロセスID (PID) =
OSがそのプロセスを識別するために付与する一意の番号。

------------------------------------------------------------------------

## 3. ユーザー作成

``` powershell
net user /add scp-user Password20250917!
```

出力例:

    パスワードが 14 文字より多くなっています。
    Windows 2000 より以前の Windows ではこのアカウントは使用できなくなります。
    この操作を続行しますか? (Y/N) [Y]: Y
    コマンドは正常に終了しました。

------------------------------------------------------------------------

## 4. パスワードを無期限にする（任意）

GUI 操作: 1. `lusrmgr.msc`を実行（ローカルユーザーとグループ管理を開く）
2. ユーザー一覧から`scp-user` をダブルクリック 3. 「全般」タブ
→「パスワードを無期限にする」にチェック

PowerShell 確認:

``` powershell
net user scp-user
```

出力例（一部）:

    アカウント有効           Yes
    アカウントの期限         無期限
    パスワード有効期間       無期限
    所属しているローカル グループ *Users

------------------------------------------------------------------------

## 5. SCP 用フォルダ作成

``` powershell
New-Item -ItemType Directory -Path C:\SCP-Share
```

### scp-user が専用で使えるフォルダにする

1.  `C:\SCP-Share` を右クリック → プロパティ\
2.  \[セキュリティ\] タブ → \[詳細設定\]\
3.  「継承の有効化を無効にする」 →
    親フォルダからの余計な権限を切り離すために実施

``` powershell
# Users グループの権限を削除
icacls C:\SCP-Share /remove "Users"

# scp-user に Modify 権限を付与
icacls C:\SCP-Share /grant scp-user:(OI)(CI)M
```
(OI) = Object Inherit  
このフォルダ内に作成される ファイルにも権限を継承する。  
(CI) = Container Inherit  
このフォルダ内に作成される サブフォルダにも権限を継承する。  
M = Modify  
読み取り・書き込み・作成・削除が可能  

### 確認方法

``` powershell
icacls C:\SCP-Share
```

期待する出力例:

    C:\SCP-Share WIN-SERVER\scp-user:(OI)(CI)(M)
                 NT AUTHORITY\SYSTEM:(OI)(CI)(F)
                 BUILTIN\Administrators:(OI)(CI)(F)

------------------------------------------------------------------------

## 6. sshd_config の編集

SCP 専用ユーザーとして `scp-user` を制御するために、`sshd_config`
を編集する。

### 設定ファイルの場所

    C:\ProgramData\ssh\sshd_config

### 編集内容（ファイル末尾に追記）

``` text
# scp-user 以外のログインを拒否
AllowUsers scp-user

# scp-user はファイル転送専用
Match User scp-user
    ForceCommand internal-sftp
    PermitTTY no
    AllowTcpForwarding no
```

#### 各設定の意味

-   **AllowUsers scp-user**\
    → scp-user 以外のログインをすべて拒否。\
-   **ForceCommand internal-sftp**\
    → ログイン時に強制的に SFTP
    サーバを実行。シェルは使えず転送専用になる。\
-   **PermitTTY no**\
    → 端末セッションを禁止。`ssh scp-user@host` は即切断。\
-   **AllowTcpForwarding no**\
    → ポートフォワーディングを禁止。SCP/SFTP のみに限定。

------------------------------------------------------------------------

## 7. sshd サービスの再起動

``` powershell
Restart-Service sshd
```

------------------------------------------------------------------------

## 8. 動作確認

### 1. scp-user でファイル転送できること

``` bash
scp test.txt scp-user@<サーバIP>:/C:/SCP-Share/
```

### 2. scp-user で SSH ログインできないこと

``` bash
ssh scp-user@<サーバIP>
```
→ 即切断されれば OK。

### 3. Administrator など他ユーザーで拒否されること

``` bash
ssh administrator@<サーバIP>
scp test.txt administrator@<サーバIP>:/C:/SCP-Share/
```
→ `Permission denied` になれば OK。

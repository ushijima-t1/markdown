# 動作確認結果

Windows Server 2022 上で OpenSSH を構成し、`scp-user` のみが SCP/SFTP
を利用できるように設定した結果の検証ログ。
wlabからSCPサーバへ確認を実施

------------------------------------------------------------------------

## 1. SCPサーバへアップロード

``` bash
scp "scp memo.txt" scp-user@192.168.122.65:/C:/SCP-Share/
```

出力例:

 [ushijima-t@wlab-kvm01 ~]$ scp "scp memo.txt" scp-user@192.168.122.65:/C:/SCP-Share/  
The authenticity of host '192.168.122.65 (192.168.122.65)' can't be established.  
ED25519 key fingerprint is SHA256:Gj/nJvwNzP2G3fPnIUkt4s5lPyAcB+z+5o32mglMRNM.  
This key is not known by any other names  
Are you sure you want to continue connecting (yes/no/[fingerprint])? yes  
Warning: Permanently added '192.168.122.65' (ED25519) to the list of known hosts.  
scp-user@192.168.122.65's password:  
scp memo.txt                                                                                                                                                                                               100%  342   244.9KB/s   00:00  

------------------------------------------------------------------------

## 2. SCPサーバからダウンロード

``` bash
scp scp-user@192.168.122.65:/C:/SCP-Share/scp\ test.txt ./scp-test.txt
```

出力例:

[ushijima-t@wlab-kvm01 ~]$ ls -l  
total 4  
-rw-r--r--. 1 ushijima-t staff 342 Sep 17 13:44 'scp memo.txt'  

 [ushijima-t@wlab-kvm01 ~]$ scp scp-user@192.168.122.65:/C:/SCP-Share/scp\ test.txt ./scp-test.txt  
scp-user@192.168.122.65's password:  
scp test.txt                                                                                                                                                                                               100%    9     3.7KB/s   00:00  

[ushijima-t@wlab-kvm01 ~]$ ls -l  
total 8  
-rw-r--r--. 1 ushijima-t staff 342 Sep 17 13:44 'scp memo.txt'  
-rw-------. 1 ushijima-t staff   9 Sep 17 13:52  scp-test.txt  

------------------------------------------------------------------------

## 3. AdministratorでSCPは拒否される

``` bash
scp test.txt administrator@192.168.122.65:/C:/SCP-Share/
```

出力例:

[ushijima-t@wlab-kvm01 ~]$ scp test.txt administrator@192.168.122.65:/C:/SCP-Share/  
administrator@192.168.122.65's password:  
administrator@192.168.122.65: Permission denied (publickey,password,keyboard-interactive).  
Connection closed  

------------------------------------------------------------------------

## 4. scp-userでSSHは拒否される

``` bash
ssh scp-user@192.168.122.65
```

出力例:

[ushijima-t@wlab-kvm01 ~]$ ssh scp-user@192.168.122.65  
scp-user@192.168.122.65's password:  
PTY allocation request failed on channel 0  
This service allows sftp connections only.  
Connection to 192.168.122.65 closed.  

------------------------------------------------------------------------

## 5. AdministratorでSSH は拒否される

``` bash
ssh administrator@192.168.122.65
```

出力例:

[ushijima-t@wlab-kvm01 ~]$ ssh administrator@192.168.122.65  
administrator@192.168.122.65's password:  
administrator@192.168.122.65: Permission denied (publickey,password,keyboard-interactive).  

------------------------------------------------------------------------

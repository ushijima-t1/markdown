## 1. 分散FWとは
- 代表例: **VMware NSX-T Data Center Distributed Firewall (DFW)**
- **仮想基盤の中で動作**するファイアウォール機能。
- 各仮想マシンの **vNIC 単位**にポリシーが適用される。
- **Inbound / Outbound** は「VMから見ての通信方向」で表現される。
  - Inbound = VM に入ってくる通信
  - Outbound = VM から外に出る通信
- **inside / outside** のゾーンを切る必要はなく、VM視点でルールを書く。
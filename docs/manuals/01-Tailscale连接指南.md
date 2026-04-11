# Tailscale 连接指南

> 本指南用于安装和配置Tailscale，连接到公司成本计算系统。

---

## 一、下载安装

### Windows 电脑：
1. 浏览器打开 https://tailscale.com/download/windows
2. 点击 **Download for Windows**
3. 运行下载的安装程序，按提示完成安装

### 手机（iPhone/Android）：
- iPhone：App Store 搜索 **Tailscale** → 安装
- Android：应用商店搜索 **Tailscale** → 安装

---

## 二、登录

1. 打开 Tailscale 应用
2. 点击 **Log in**
3. 使用分配给你的账号登录
4. 登录后显示等待页面是正常的，请通知CC审批

---

## 三、审批通过后

CC审批通过后，Tailscale状态会自动变为 **Connected**。

此时浏览器输入以下地址即可访问成本计算系统：

```
http://100.102.56.99:8502/
```

建议收藏此地址，方便下次使用。

---

## 四、日常使用

- Tailscale 默认开机自启动，无需手动操作
- 如果状态显示 Not connected，点击 **Connect** 重新连接即可

---

## 五、常见问题

| 问题 | 解决方法 |
|------|---------|
| 登录后一直等待 | 已通知CC，等待审批即可 |
| 审批后无法访问网页 | 确认Tailscale显示Connected |
| 页面打不开 | 确认地址是 http://100.102.56.99:8502/（http不是https） |
| Tailscale断开 | 打开Tailscale点击Connect重连 |

---

*如有问题请联系CC*

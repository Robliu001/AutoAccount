# 与GIT仓库远程通信
## 生成key
    ssh -keygen -t rsa。public key存储路径/c/Users/bqliu/.ssh/id_rsa.pub.
    在git上Add new ssh key，把id_rsa.pub内容添加进去
## 测试连接
[**ssh -T git@github.com**]<br>
第一次会有如下显示：<br>
The authenticity of host 'github.com (13.250.177.223)' can't be established.
RSA key fingerprint is SHA256:nThbg6kXUpJWGl7E1IGOCspRomTxdCARLviKw6E5SY8.
Are you sure you want to continue connecting (yes/no)?
输入yes，第二次就不会有了。
## 设置用户名和Email
**git config --global user.name "Qiang"**<br>
**git config --global user.email "liuboqiang@126.com"**<br>
#要和github上的一致，主页上就会显示出添加的绿色格子
## 生成与github上某个仓库对应的远程库名称
    git remote add ying git@github.com:Robliu001/AutoAccount.git
## 获取远程库
    git remote -v
## 强制推送
    git push -u ying master -f

    在github上创建项目，然后本地git init
    然后没有git pull -f --all
    然后git add .  | git commit -am "init"
    导致github上的版本里有readme文件和本地版本冲突，下面给出冲突原因：
    \[master][~/Downloads/ios] git push -u origin master
    Username for 'https://github.com': shiren1118
    Password for 'https://shiren1118@github.com':
    To https://github.com/shiren1118/iOS_code_agile.git
     ! [rejected]        master -> master (non-fast-forward)
    error: failed to push some refs to 'https://github.com/shiren1118/iOS_code_agile.git'
    hint: Updates were rejected because the tip of your current branch is behind
    hint: its remote counterpart. Merge the remote changes (e.g. 'git pull')
    hint: before pushing again.
    hint: See the 'Note about fast-forwards' in 'git push --help' for details.

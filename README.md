# outlook_bulk_action
bulk outlook email 's  delete / move

# install 

cpan Mail::Outlook Win32::OLE Date::Calc

# bulk delete email

perl delete_outlook_mail.pl -m "somemailbox/miscfolder/log" -f "abc@xxx.com" -t "some log" -b "all ok" -D 

    -m mailbox path, for example, "somemailbox/miscfolder"；指定邮件文件夹路径，somemailbox为邮箱名，miscfolder/log为二层子邮件夹
    -f mail from, 发件人，例如"abc@xxx.com"，也可以指定多个，例如 "SOMEGRP,someuser@xxx.com"
    -t mail subject, 邮件标题，例如 "some log"
    -b mail message, 邮件正文，例如"all ok"
    -d select mail arrived n days ago, 邮件发送时间距离现在有多少天，默认为21天
    -D not check mail arrived time, 不管邮件发送时间是什么时候，满足指定条件就删除(覆盖-d参数)

# bulk move email
```
perl move_outlook_mail.pl [src mailbox path] [dest mailbox path]

perl move_outlook_mail.pl srcmailbox/工作  dstmailbox/存档 
```

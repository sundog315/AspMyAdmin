# AspMyAdmin
AspMyAdmin is a simple ASP Scripts,Use to Manage small database like an ACCESS database or a small SQLserver database.
this Script was Writed in 2001 ,When i was a student.

ASPmyadmin说明文件
Version: beta0.9

    这是一套数据库管理软件。适用于小型用户。（譬如：建立自己的站点，用本系统来管理数据库。）作者也是一个ASP
爱好者，一直都想找一个方便而且便于使用的数据库管理软件。可是，这方面的软件要么很大，用不着。要么就是E文的，看
着太累。于是作者便决定自己编制一个软件。
    
    使用说明：
    1.用户需要更改"config.asp"中的连接字段，与自己的数据库进行连接。因为ADOLess和ODBC不支持OPENSHEMA的Primarykey
属性，因此强烈建议使用JET。如果用ADOless和ODBC本系统将不能识别主键。

    2.第一次运行本系统将会在所指向的数据库中插入一个名为AspMyAdmin_user的表。
   
    注意事项：
    1.由于本系统以主键来区分不同的纪录，所以表中必须有主键，否则会发生错误。
    3.请在第一次使用前更改密码加密密钥，更改"function.asp"中的"sBASE_64_CHARACTERS"字串使加密后的密码不易被破解。
（字串中不能出现重复的字母，而且字段长度不能改变。(64位)）!!!!!!请小心使用，更改后请重新登陆请参考"更改密钥说明.txt"

   
    由于是beta版本，可能会产生不可遇见的错误。造成的任何损失均与本人无关。
    如果您觉得这个系统还不错，请来信鼓励我。谢谢！
    由于作者是一名在校学生，来信可能不能迅速回答，请见谅。我一定尽可能快的回答大家的问题。
    
    欢迎大家交换意见：sundog315@foxmail.xom

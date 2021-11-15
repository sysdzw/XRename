# XRename
文件文件夹超级重命名工具
XRename又称文件文件夹超级重命名工具，可以帮助你快速的将一批文件或者文件夹根据指定的要求重新命名，比如将指定目录下所有文件的文件名中“卷”替换成“册”。此工具最大的特色是可以允许高级用户使用正则表达式设置自己的规则，要处理的文件范围也可以用正则表达式来限定，这样基本就万能了。下面来讲讲XRename的命令及用法吧。


二级命令

1.替换文件名中的字符，语法如下：
replace -dir directory -string string1 -(new|newstring|replacewith) string2 [-type(file|dir|all)[:string3]] [-subdir(yes|no)] [-ignorecase(yes|no)] [-log(yes|no)]

功能说明：将某个目录中的文件或文件夹的名称中的字符按指定规则替换，regexp1和regexp2表示可以使用正则表达式。

参数说明：
具体的参数值建议都加上双引号，因为如果参数值里面有空格的话会影响程序的判断。用正则表达式的话除外，因为它已经用//表示了。
-dir        要处理的目录，也可以写作-path。

-string        要替换的字符串。这里可以用正则表达式，格式为“/regexp/img”，和js脚本中的设置一样，注意它外围不能再加双引号，否则只会被当做普通字符串处理。正则表达式的匹配属性可以在第二个/后面控制，忽略大小写用i，多行匹配用m，匹配所有项用g，因为文件名没有换行的，所以加不加m都是一样的。正则表达式默认匹配属性为“区分大小写”和“非全局匹配”。

-new    替换后的字符串，还可以写作-newstring和-replacewith。如果前面的-string用的正则表达式那么这里可以用“$1”或“$2”这样的分组捕获内容，否则只会被当做普通字符串处理。

-type        要处理的对象的类型，这里共有三种情况。即file（文件），dir（文件夹）以及all（包含前面两者）。默认为file，也就是只处理文件，这个参数后面还可以加上“:”然后指定处理范围。这里可以用正则表达式也可以用普通字符。普通字符的话就是固定一个字符串或者匹配字符串，和windows匹配方式兼容，例如*.txt就是指处理所有txt文件，?就表示单个字符。如果要用正则表达式那么和-string参数使用正则表达式情况的要求一样的。

-subdir      是否需要处理子目录。yes为处理，并且会递归访问子目录，no则不处理子目录。默认为no，表示只处理当前文件夹下的所有文件或者文件夹

-ignorecase    是否忽略字母大小写。yes为忽略，即不区分字母大小写，no则区分。默认为yes，这个在-string使用普通字符串时会用到，如果是用正则表达式的话会由/后面的标记i来决定。

-log        是否输出处理日志，文件名为XRename.log。yes为输出，no则不输出，默认为no，表示不生产log文件。另外如果XRename在处理时发生错误的情况下无论是否指定-log这个参数都会生成一个名为XRename_err.log的文件。

应用范例：
(1)将"c:\movie\"下所有文件的文件名中的"老友记"替换为"friends"
XRename replace -dir "c:\movie\" -string "老友记" -replacewith"friends"

(2)将"c:\movie\"下所有文件的文件名中的空格替换为下划线，并且生成log
XRename replace -dir "c:\movie\" -string " " -replacewith "_" -log yes

(3)将"c:\movie\"下所有以wma为后缀名的文件替换为rmvb后缀名。
XRename replace -dir "c:\movie\" -string "wma" -replacewith "rmvb"

上面的方法可能不保险，因为必须最后是wma的才替换，可以使用正则表达式精确处理：
XRename replace -dir "c:\movie\" -string /(.*?)wma$/ig -replacewith "$1rmvb" 或：
XRename replace -dir "c:\movie\" -string /wma$/ig -replacewith "rmvb"

如果需要进一步缩小范围指定处理wma文件，那么用下面方法：
XRename replace -dir "c:\movie\" -string /wma$/ig -replacewith "rmvb" -type file:"*.wma" 或
XRename replace -dir "c:\movie\"-string /wma$/ig -replacewith "rmvb" -type file:"/.*\.wma/ig"


2.删除文件名中的字符，语法：
delete -dir directory -string string1 [-type (file|dir|all)[:string3]] [-subdir (yes|no)] [-ignorecase (yes|no)] [-log(yes|no)]

功能说明：将某个目录中的文件或文件夹的名称中的字符按指定规则的删除。此命令实际可用replace命令代替，即替换为空。

参数说明：参考replace功能的参数说明部分。

应用范例：
(1)将"c:\movie\"下所有文件的文件名中的"book"删除
XRename delete -dir "c:\movie\" -string"book"

(2)将"c:\inet\"下所有文件的文件名中的"["和"]"删除，这个应用很典型，例如从ie临时文件夹拷贝出来的文件基本都会带有字符[1]和[2]字样的

XRename delete -dir "c:\inet\" -string "/\[|\]/ig"


如果要直接把[1]或[2]删除的话，可以用下面的方法，不过可能会引起冲突

XRename delete -dir "c:\inet\" -string "/\[\d+\]/ig"




3.列出文件名，语法：
listfile -dir directory -string string1 [-type(file|dir|all)[:string3]] [-subdir(yes|no)] [-ignorecase (yes|no)] [-output path]

功能说明：导出某个目录下符合指定规则的文件或文件夹的名称列表。

参数说明：参考replace功能的参数说明部分。其中-output为导出的列表保存的路径，默认为指定目录下的XRename_list.txt文件。

应用范例：
(1)列出"c:\movie\"下所有文件的文件名含有"经典"的文件
XRename listfile -dir "c:\movie\" -string "经典"

(2)列出"c:\movie\"下所有文件的文件名以"经典"二字开头并且以CD1结尾（忽略后缀名）的文件，并将内容导出到"c:\classicMovie.txt"
XRename listfile -dir "c:\movie\" -string /^经典.+?CD1(\.[^\.]*)?/ig -output "c:\classicMovie.txt"


4.删除文件，语法：
delfile -dir directory -string string1 [-type (file|dir|all)[:string3]] [-subdir (yes|no)] [-ignorecase (yes|no)] [-log (yes|no)]

功能说明：删除某个目录下符合指定规则的文件或文件夹。

参数说明：参考replace功能的参数说明部分。

应用范例：
(1)删除"c:\movie\"下所有文件名含有"苍井空"的文件
XRename delfile -dir "c:\movie\" -string "苍井空"

(2)删除"c:\test\"下所有目录名为数字的目录，包含子目录。subdir 表示是否包含子目录
XRename delfile -dir "c:\test\" -string "^\d+$"-type dir -subdir yes


5.UTF8类型的解码，语法：
utf8rename -dir directory [-type (file|dir|all)[:string3]] [-subdir (yes|no)] [-ignorecase (yes|no)] [-log (yes|no)]

功能说明：将文件名用UTF8编码的文件进行文件名解码，主要应用于对从IE临时文件夹拷贝的文件重命名。

 

应用范例：

XRename utf8rename -dir "c:\movie\"


6.其他待补充。


另外说明下：
默认要替换的字符即-string后面的实际都是当做正则表达式的，所以某些字符（正则表达式的元字符，也就是关键字符）是需要转义的，假设需要将“.”替换成"-"，因为那两个字符在正则表达式中都表示特殊的意思，如果你要替换的字符就是指“.”的话那么需要写成"\."来转义，这个实际是正则表达式的知识了。 还有一个需要特别说明的是，由于所有参数基本都需要用半角双引号引起来，但是你需要替换的字符就是含有双引号怎么办呢？XRename中的方案是用\转义。例如将文件名中双引号删除掉，那么用XRename delete -dir "c:\movie\" -string "\""
————————————————
版权声明：本文为CSDN博主「无·法」的原创文章，遵循CC 4.0 BY-SA版权协议，转载请附上原文出处链接及本声明。
原文链接：https://blog.csdn.net/sysdzw/article/details/6198257

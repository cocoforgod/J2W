# J2W  Java 使用 POI 3.17根据Word 模板替换、操作书签
由于项目的需求，需要对大量的word文档进行处理。

查找了大量的文档发现很多的博客对这个进行了介绍，主要有2种方案做处理，jacob 和poi。但是现在的服务器基本上是部署在Linux上，所以jacob基本上是不可行的。所以呢，主要是使用poi来进行这些操作。

       Apache poi的hwpf模块是专门用来对word doc文件进行读写操作的。在hwpf里面我们使用HWPFDocument来表示一个word doc文档。在HWPFDocument里面有这么几个概念：
 Range：它表示一个范围，这个范围可以是整个文档，也可以是里面的某一小节（Section），也可以是某一个段落（Paragraph），还可以是拥有共同属性的一段文本（CharacterRun）。

 Section：word文档的一个小节，一个word文档可以由多个小节构成。

 Paragraph：word文档的一个段落，一个小节可以由多个段落构成。

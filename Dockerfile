#指定基础镜像，在其上进行定制`
FROM java:8
#SpringBoot 项目默认使用 `/tmp'目录作为工作目录，省去了复制文件的麻烦
#这里的 /tmp 目录就会在运行时自动挂载为匿名卷，任何向 /data 中写入的信息都不会记录进容器存储层
VOLUME /tmp
#复制上下文目录下的打包好的jar文件并重命名到容器里
ADD /target/parsegameexcel-0.0.1-SNAPSHOT.jar parsegameexcel.jar
#声明运行时容器提供服务端口，这只是一个声明，在运行时并不会因为这个声明应用就会开启这个端口的服务
EXPOSE 8080
#指定容器启动程序及参数"
ENTRYPOINT ["java","-jar","parsegameexcel.jar"]
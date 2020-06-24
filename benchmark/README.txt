Although POI is free, but is slowest, and the most critical problem of POI is memory, this is the reason why I just choose 100,000*30 cells(GcExcel Java can run smoothly even with 1,000,000 *30 cells) , even in such amount of data, you must configure JVM min heap size to 4G to make POI run success, or else it will throw OutOfMemory exception. below is the JVM configuration in build.gradle 

applicationDefaultJvmArgs = ["-Xms4096m", "-Xmx8192m"]
 

To run this project, use below commands line

on mac: ./gradlew run --args="double" //args can be double, string, date, formula, bigfile
on windows: gradlew run --args="double" //args can be double, string, date, formula, bigfile
Note that
1. the first time of running will be very slow, because it needs to download gradle and all the dependent packages.
2. The calculating of used memory may not be accurate, it depends on the state of current JVM Heap, so you'd better run several times.
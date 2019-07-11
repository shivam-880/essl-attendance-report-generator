# essl-attendance-report-generator
Generates attendance report in Microsoft excel from biometric attendance logs.

## Dev
```
$ cd essl-attendance-report-generator
$ sbt> universal:stage 

$ cd ./target/universal/stage/logs
$ tail -f stdout.log essl-attendance-report-generator.log
```

## Build Package
```
$ cd essl-attendance-report-generator
$ sbt> universal:packageBin

$ cd ./logs
$ tail -f stdout.log essl-attendance-report-generator.log
```

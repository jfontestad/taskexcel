# utility that retrieves planner tasks for the TSTA team
# computes effort points
# and saves them to two CSV:
# all tasks contains all tasks
# active tasks contains tasks that are not completed, nor waiting nor problems
# csv can be downloaded using the file/open menu above

#configuration parameters
#gather_descriptions indicates if the descriptions of the active tasks shall be downloaded
#in 15% of cases gathering descriptions leads to an error message
print("start batch")
for (count in seq(32)) {
    print(paste("start iteration ", count, sep=""))
    load(paste("/home/docker/myEnvironmentReadAllTasksExcel.RData",sep=""))

    #if uncommented redirects stdout and stderr to a file
    #logFile = file(paste("/home/docker/myEnvironmentReadAllTasksExcel.log",sep=""))
    #sink(logFile, type = c("output"))
    #sink(logFile, type = c("message"))

    print(paste("*** Read all tasks starting at: ", Sys.time()))

    gather_descriptions = FALSE

    source(paste("/home/docker/ReadAllTasksExcel.R",sep=""))

    save.image(file=paste("/home/docker/myEnvironmentReadAllTasksExcel.RData",sep=""))

    print("image saved, job completed")

    Sys.sleep(900) #15 minutes intervals
}
print(paste("*** Read all tasks closing at: ", Sys.time()))


#Read All Tasks Excel
# version 20200809 0918
# utility that retrieves planner tasks for the TSTA team
# computes effort points
# and saves them to two excel:
# all tasks contains all tasks
# active tasks contains tasks that are not completed, nor waiting nor problems
# csv can be downloaded using the file/open menu above

#known issues
# column "points" is displayed in excel as text
# having the link in the last column is not comfortable


#configuration parameters
#gather_descriptions indicates if the descriptions of the active tasks shall be downloaded
#in 15% of cases gathering descriptions leads to an error message
gather_descriptions = FALSE

#configuration: list below the planner plans and the corresponding team names to be retrieved
plans_and_groups = list(c('Tech Invoices 2020','Tech Invoices'), 
                        c('Transformation Assurance','Transformation Assurance'),
                        c('Tech Specific Contracts 2020','Tech Specific Contract'),
                        c('Tech Forecast','Tech Forecast'),
                        c('Tech Order Forms 2020','Tech Order Forms'),
                        c('BVTECH Orders','BVTECH Orders'),
                        c('Tech DS 2020','Tech Delivery Slips'),
                        c('TECH EC Inter-institutional', 'Tech Order Forms'),
                        c('TS TA KAIZEN', 'LEAN'),
                        c('Tech Management Follow up', 'Tech Management Follow up'))

#test
#plans_and_groups = list(c('TECH EC Inter-institutional', 'Tech Order Forms'))

# install package AzureGraph if not yet installed
library(AzureGraph)
library(jsonlite)
library(tidyverse)
library(bizdays)
library(httr)
library(functional)
library(gtools)
library (stringi)
library(httpuv)

### logon - START
if (!exists("me")) {
  gr=create_graph_login(auth_type="device_code")
  me <- gr$get_user()
}
### logon - END

# this has been used once to retrieve the plan ids
# Tech Invoices 24b46af2-2dff-4eb5-aea2-c7218dc6151b ; Ghcj_Mhvsk6TmMxVcTTIBpYAF7IO
# Transformation Assurance 58398e3c-899e-431f-99d4-207f15012489 ; 1W6x6L_52Eyp4S_bEvZdD5YADszD
# Tech SC 13b63b22-be1e-45e8-b1a5-51be7cd8ac5e ; RjBjjQ99bEWOT20KuR4sG5YADKhq
# Tech Forecast 35da596a-ea9d-491e-ab88-e5f5896e338a ; o4ZH9L5rQES_Zc0EoOfED5YAAiDB
# Tech OF 53227348-d85e-4111-abe8-00c7f31405a9 ; HivMTTcwNkyWGDgAuOQtc5YAFx9O
# BVTECH 7fee428e-b61e-4fe0-ad89-1cdbee8eb7f7 ; OBLthOEvvES_1zpfiT6W2JYAAjyE
# planner_plans = call_graph_url(me$token,"https://graph.microsoft.com/v1.0/groups/7fee428e-b61e-4fe0-ad89-1cdbee8eb7f7/planner/plans")
# planner_plans

### retrieve data from planner - START
my_teams = call_graph_url(me$token, "https://graph.microsoft.com/beta/me/joinedTeams")


retrieve_plan_id = function (plan_and_group) {
  plan_name = plan_and_group[[1]]
  group_name = plan_and_group[[2]]
  group_id = my_teams$value[unlist(map(my_teams$value, function(x) x$displayName == group_name))][[1]]$id
  plan_data = call_graph_url(me$token,paste("https://graph.microsoft.com/v1.0/groups/",group_id,"/planner/plans",sep=""))
  plan_id = plan_data$value[unlist(map(plan_data$value, function(x) x$title == plan_name))][[1]]$id
  plan_id
}

planner_plan_ids = map(plans_and_groups, function(x) c(x[1], retrieve_plan_id(x)))

#planner_plan_ids = list(c('Tech Invoices 2020','Ghcj_Mhvsk6TmMxVcTTIBpYAF7IO'), 
#                        c('Transformation Assurance','1W6x6L_52Eyp4S_bEvZdD5YADszD'),
#                        c('Tech Specific Contracts 2020','RjBjjQ99bEWOT20KuR4sG5YADKhq'),
#                        c('Tech Forecast','o4ZH9L5rQES_Zc0EoOfED5YAAiDB'),
#                        c('Tech Order Forms 2020','HivMTTcwNkyWGDgAuOQtc5YAFx9O'),
#                        c('BVTECH Orders','OBLthOEvvES_1zpfiT6W2JYAAjyE'),
#                        c('Tech DS 2020','JNiaOP_GyUyMNbeyksSKlJYAHW1d'))

#this could be used to determine the plan ids 
#using dynamic lookups
#plan_names = c("Tech Delivery Slips")

#same as previous cell
#my_teams$value[lapply(my_teams$value, function(x) x$displayName) %in% plan_names]

user_list = list(user_names = c('AG', 'MN', 'DP', 'ZV', 'BG', 'unassigned'),
                 user_ids = c('3963a97a-43e7-43e7-b71a-a710b7bbc4fa','7900481f-40d8-4e73-868b-6db453b53aa4','0db6b33a-4615-43af-85de-cab76eb2628c','37a32389-1aac-481c-a1fc-e80ab2e5aa34', 'ab4bbfc2-6d90-44f6-b832-86ae7c6464a2', 'NA'))
user_tib = as_tibble(user_list)

priority_list = list(priority_names = c('Low', 'Medium', 'Important', 'Urgent'),
                     priority_ids = c('9','5','3','1'))
priority_tib = as_tibble(priority_list)

category_list = list(categories = c("Problem", "Huge Effort", "High Effort", "ISO ok", "New Message", "Waiting"),
                     category_ids = sapply(seq(6), function(x) paste("category", x, sep="")))
category_tib = as_tibble(category_list)

det_effort = function (category_list) {
  #include only effort categories
  if (length(category_list[category_list %in% 'category2'] ) == 1) {
    'category2'
  } else {
    if (length(category_list[category_list %in% 'category3'] ) == 1) {
      'category3'
    } else {
      'NA'
    }
  }
}

det_problem = function (category_list) {
  if (length(category_list[category_list %in% 'category1'] ) == 1) {
    'category1'
  } else {
    'NA'
  }  
}

det_waiting = function (category_list) {
  if (length(category_list[category_list %in% 'category6'] ) == 1) {
    'category6'
  } else {
    'NA'
  } 
}

#det_new_message
det_new_message = function (category_list) {
  if (length(category_list[category_list %in% 'category5'] ) == 1) {
    'category5'
  } else {
    'NA'
  }  
}

ensure_value = function (uncertain_val) {
  if (is.null(uncertain_val)) {
    'na'
  } else {
    uncertain_val 
  }
}

cut_date_time = function (date_time_string) {
  if (nchar(date_time_string) >= 10) {
    temp_date_time = substring(date_time_string, 1, 10)
  } else {
    temp_date_time = date_time_string
  }    
#  format(strptime(temp_date_time, format="%Y-%m-%d"), format="%d/%m/%Y")
  strptime(temp_date_time, format="%Y-%m-%d")
}

#det_effort_points
# calls det_effort_points_single for a vector of categories
det_effort_points = function (effort_category_list) {
  unlist(lapply(effort_category_list, det_effort_points_single))
}

#def_effort_points_single
# reads the categories to determine the effort for one 
# category entry
det_effort_points_single = function (effort_category) {
  switch(
    effort_category,
    "Huge Effort" = 4,
    "High Effort" = 2,
    "Medium Effort" = 0,
    "Low Effort" = 0, #low effort is not used anymore
    "NA" = 0)
}

# function that reads all tasks for a plan (identified by its id)
#  and extracts selected columns
#  todo: insert a cache so that for already loaded plan_name we use a cached table and not read every time
extract_task_tbl = function(plan_name_and_id) {
  plan_name = plan_name_and_id[1]
  plan_id = plan_name_and_id[2]
  inv_tasks = call_graph_url(me$token,paste("https://graph.microsoft.com/beta/planner/plans/",plan_id,"/tasks", sep=""))
  num_rows = length(inv_tasks$value)
  #extract assignments (that are nested)
  inv_tasks_assignments = lapply(inv_tasks$value, function(x) x$assignments)
  inv_tasks_assignments_ids = lapply(inv_tasks_assignments, names)
  inv_tasks_assignments_ids = lapply(inv_tasks_assignments_ids, function (xx) if (length(xx) == 0) "NA" else xx[[1]])  
  inv_tasks_categories = lapply(inv_tasks$value, function(x) names(x$appliedCategories))
  inv_tasks_effort = lapply(inv_tasks_categories, function(x) det_effort(x))
  inv_tasks_problem = lapply(inv_tasks_categories, function(x) det_problem(x))
  inv_tasks_waiting = lapply(inv_tasks_categories, function(x) det_waiting(x))
  inv_tasks_new_message = lapply(inv_tasks_categories, function(x) det_new_message(x))
  #compose the initial tibble that contains only the plan name repeated for each row
  tib_inv_tasks = as_tibble(list(plan=rep(plan_name, num_rows)))
  # compose the lists and clean the values (they are all lists and need sapply fun=paste to become char)
  tib_inv_tasks = tib_inv_tasks %>%
    mutate(title=sapply(lapply(inv_tasks$value, function(x) x$title), FUN=paste)) %>%
    mutate(bucketId = sapply(lapply(inv_tasks$value, function(x) x$bucketId), FUN=paste)) %>%
    mutate(createdDateTime = sapply(lapply(inv_tasks$value, function(x) cut_date_time(x$createdDateTime)), FUN=paste)) %>% 
    mutate(startDateTime = sapply(lapply(inv_tasks$value, function(x) cut_date_time(ensure_value(x$startDateTime))), FUN=paste)) %>%
    mutate(dueDateTime = sapply(lapply(inv_tasks$value, function(x) cut_date_time(ensure_value(x$dueDateTime))), FUN=paste)) %>%
    mutate(completedDateTime = sapply(lapply(inv_tasks$value, function(x) cut_date_time(ensure_value(x$completedDateTime))), FUN=paste)) %>%
    mutate(id=sapply(lapply(inv_tasks$value, function(x) x$id), FUN=paste)) %>%
    mutate(priority=sapply(lapply(inv_tasks$value, function(x) x$priority), FUN=paste)) %>%
    mutate(assignments = unlist(inv_tasks_assignments_ids)) %>%
    mutate(effort = sapply(inv_tasks_effort, FUN=paste)) %>%
    mutate(problem = sapply(inv_tasks_problem, FUN=paste)) %>%
    mutate(waiting = sapply(inv_tasks_waiting, FUN=paste)) %>%
    mutate(new_message = sapply(inv_tasks_new_message, FUN=paste))
  
  tib_inv_tasks
}

#repeat gathering tasks for each plan in the list planner_plan_ids
integrated_tbl = tibble()
integrated_tbl = map(planner_plan_ids, function(x) extract_task_tbl(x))
integrated_tbl = bind_rows(integrated_tbl)

# list the buckets by ID
extract_buckets = function (plan_name_and_id) {
  plan_name = plan_name_and_id[1]
  plan_id = plan_name_and_id[2]
  buckets = call_graph_url(me$token,paste("https://graph.microsoft.com/v1.0/planner/plans/",plan_id,"/buckets", sep=""))
  num_rows = length(buckets$value)
  tib_buckets = as_tibble(list(b_plan = rep(plan_name, num_rows)))
  tib_buckets = tib_buckets %>%
    mutate(b_name=sapply(lapply(buckets$value, function(x) x$name), FUN=paste)) %>%
    mutate(b_id=sapply(lapply(buckets$value, function(x) x$id), FUN=paste))
}

#gather buckets for each plan
integrated_buckets = tibble()
integrated_buckets = map(planner_plan_ids, function(x) extract_buckets(x))
integrated_buckets = bind_rows(integrated_buckets)

#varibles that show progress in description download
num_tasks = 0
cur_task = 0

# add description
getDescription <- function (taskId) {
  task_url = paste ("https://graph.microsoft.com/v1.0/planner/tasks/",toString (taskId),"/details",sep="")
  task_det = call_graph_url(me$token,task_url)
  json_task_det <- fromJSON(toJSON(task_det, flatten=TRUE, pretty=TRUE))
  cur_task <<- cur_task + 1
  print(paste("Retrieving description",cur_task, "out of",num_tasks))
  #tib_task_det = as_tibble(json_task_det$value[,c("title", "bucketId","createdDateTime","dueDateTime", "completedDateTime", "id")]) 
  json_task_det$description
}

getDescVector <- function (ids) {
  sapply(replace_na(lapply(ids, getDescription)), FUN=paste)
}

desc_tbl <- integrated_tbl %>% 
  filter(completedDateTime == 'NA')


num_tasks = length(desc_tbl$id)

if (gather_descriptions) {
  desc_tbl <- desc_tbl %>%
    mutate(description=getDescVector(id)) 
} else {
# when we do not want a description to be extracted
  desc_tbl <- desc_tbl %>%
    mutate(description="Description was not extracted") 
}

non_desc_tbl <- integrated_tbl %>%
  filter(completedDateTime != 'NA')

all_tasks_tbl <- bind_rows(desc_tbl, non_desc_tbl)

#for testing purposes
#integrated_tbl

### retrieve the points from a csv stored in sharepoint - START
#search for my teams
myTeams = call_graph_url(me$token,"https://graph.microsoft.com/v1.0/me/joinedTeams")
TSTA_team = myTeams$value[lapply(myTeams$value, function(x) x$displayName) %in% "Transformation Assurance" == TRUE]
TSTA_team_id = TSTA_team[[1]]$id

# id of the team transformation assurance 58398e3c-899e-431f-99d4-207f15012489
#call_graph_url(me$token,paste("https://graph.microsoft.com/beta/groups/58398e3c-899e-431f-99d4-207f15012489/drive"))
group_drive = call_graph_url(me$token,paste("https://graph.microsoft.com/beta/groups/",TSTA_team_id,"/drive", sep=""))
group_drive_id = group_drive$id

search_csv = call_graph_url(me$token,
                            paste("https://graph.microsoft.com/v1.0/drives/",group_drive_id,"/root/search(q='{TSTAServicesTaskEfforts}')", sep="")) 
#    "https://graph.microsoft.com/v1.0/drives/b!vYX2rXwiGkeH1tDi6Gcg3GzwuQ5ZoQdDjsFgvnewpev2F8IiPURjSriJ-blxYH7a/root/search(q='{TSTA_Services_Task_Effort.csv}')")  
csv_id = search_csv$value[[1]]$id


#id of csv file 013HY5IZRMXAPN7H5FXRCZYGQI5UGSIZMU
#csv_id = "013HY5IZRMXAPN7H5FXRCZYGQI5UGSIZMU"

#read the file
csv_param = call_graph_url(me$token,paste("https://graph.microsoft.com/v1.0/drives/",group_drive_id,"/items/",csv_id,"/content", sep=""))

# convert a sequence of unicode hex to a vector or charachters 

csvchars = chr(csv_param)

# convert a vector of charachters to a string
csvtable = paste(csvchars, collapse="")
csvLongStr = toString (csvtable)
csv_fine_str = stri_enc_toutf8(stri_trans_nfd(csvLongStr)) #this final command eliminates the special charachters

# create a tibble from the csv data read from the file and converted using the previous snippets
tib_param = read_csv(csv_fine_str) %>%
  rename(Param_Plan = Plan,
         Param_Bucket = 'Bucket Name')
### retrieve the points from a csv stored in sharepoint - END

### prepare the final dataset for export - START
# the next line was eliminated because it is inserted incorrectly by the api of add column
#   mutate(url = paste("=hyperlink(\"https://tasks.office.com/EFSA815.onmicrosoft.com/en-gb/Home/Task/",id,"\",\"Link\")",sep = "")) %>%
#   mutate(StepDeadline = "=iferror(IF(offset(INDIRECT(ADDRESS(ROW(), COLUMN())),0,-1)<>\"\", NETWORKDAYS.INTL(offset(INDIRECT(ADDRESS(ROW(), COLUMN())),0,-1),offset(INDIRECT(ADDRESS(ROW(), COLUMN())),0,1))*VLOOKUP(offset(INDIRECT(ADDRESS(ROW(), COLUMN())),0,11),models!$A$1:$D$200,4,FALSE)+offset(INDIRECT(ADDRESS(ROW(), COLUMN())),0,-1),\"\"),\"\")") %>%
tasks_with_bucket_names = all_tasks_tbl %>% left_join(integrated_buckets, by=c("bucketId" = "b_id")) %>%
  left_join(user_tib, by=c('assignments' = 'user_ids')) %>%
  left_join(priority_tib, by=c('priority' = 'priority_ids')) %>%
  left_join(category_tib, by=c('effort' = 'category_ids')) %>%
  left_join(category_tib, by=c('problem' = 'category_ids')) %>%
  left_join(category_tib, by=c('waiting' = 'category_ids')) %>%
  left_join(category_tib, by=c('new_message' = 'category_ids')) %>%
  left_join(tib_param, by=c('plan' = 'Param_Plan', 'b_name' = 'Param_Bucket')) %>%
  mutate(eff_points = Points + det_effort_points(categories.x)) %>%
  mutate(startDateTime = sapply(lapply(startDateTime, function(x) {if (x == "NA") {""} else {x}}), FUN=paste))  %>%   #mutate(dueDateTime = format(dueDateTime, format="%Y-%m-%d")) %>%
  mutate(dueDateTime = sapply(lapply(dueDateTime, function(x) {if (x == "NA") {""} else {x}}), FUN=paste))  %>%   #mutate(dueDateTime = format(dueDateTime, format="%Y-%m-%d")) %>%
  mutate(completedDateTime = sapply(lapply(completedDateTime, function(x) {if (x == "NA") {""} else {x}}), FUN=paste))  %>%   #mutate(dueDateTime = format(dueDateTime, format="%Y-%m-%d")) %>%
  mutate(eff_points = Points + det_effort_points(categories.x)) %>%
  mutate(StepDeadline = "=IFERROR(IF(OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())),0,-1)<>\"\", WORKDAY.INTL(OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())),0,-1),ROUNDDOWN(NETWORKDAYS.INTL(OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())),0,-1),OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())),0,1))*VLOOKUP(OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())),0,11),models!$A$1:$D$200,4,FALSE),0)-1),\"\"),\"\")") %>%
  mutate(url = paste("https://tasks.office.com/EFSA815.onmicrosoft.com/en-gb/Home/Task/", id, sep="")) %>%
  mutate(search = "=IFERROR(CONCATENATE(\"https://www.office.com/search?auth=2&q=\",MID(OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())),0,-17),FIND(\"(TX\",OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())),0,-17)),11)),\"\")") %>%   #  mutate(search = "=IFERROR(CONCATENATE(\"https://www.office.com/search?auth=2&q=\",MID(OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())),0,-17),FIND(\"(TX\",OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())),0,-17)),11)),\"") %>%
  mutate(reminder = "=IFERROR(MID(OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())),0,-18),FIND(\"RMD\",OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())),0,-18))+3,10),\"\")") %>%   
  mutate(txid = "=IFERROR(MID(OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())),0,-19),FIND(\"(TX\",OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())),0,-19)),11),\"\")") %>%   
  mutate(late = "=IFERROR(IF(OR(OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())),0,-14)>OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())),0,-15),AND(OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())),0,-14)=\"\", OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())),0,-15)<NOW())),\"yes\", \"no\"), \"\")") %>%   
  mutate(link = paste("=hyperlink(offset(indirect(address(row(), column())), 0, 17),\"Link\")",sep = "")) %>%
  mutate(ModelKey = paste(plan, b_name, sep = "")) %>%
  select(plan, title, b_name, createdDateTime, startDateTime, StepDeadline, dueDateTime, completedDateTime, priority_names, user_names, categories.x, categories.y, categories.x.x, categories.y.y, eff_points, description, ModelKey, url,search,reminder, txid, late, link) %>%
  rename(Plan = plan,
         Title = title,
         Bucket = b_name,
         Created = createdDateTime,
         Start = startDateTime, 
         StepDue = StepDeadline,
         Due = dueDateTime,
         Completed = completedDateTime,
         Priority = priority_names,
         Assigned = user_names,
         Effort = categories.x,
         Problem = categories.y,
         Waiting = categories.x.x,
         NewMessage = categories.y.y,
         Points = eff_points,
         Description = description)

### prepare the final dataset for export - END

### export the data to a file first and then to sharepoint - START


empty_table = function(excel_file_id, task_tbl) {
  session = call_graph_url(me$token,
                           paste("https://graph.microsoft.com/v1.0/drives/",
                                 group_drive_id,
                                 "/items/",
                                 excel_file_id, 
                                 "/workbook/createSession",
                                 sep = ""),
                           body=toJSON (list(persistChanges = TRUE), auto_unbox = TRUE), 
                           http_verb=c ("POST"),
                           encode = "raw")  
  
  
  tables = call_graph_url(me$token,
                          paste("https://graph.microsoft.com/v1.0/drives/",
                                group_drive_id,
                                "/items/",
                                excel_file_id, 
                                "/workbook/tables",
                                sep="")) 
  
  if(length(tables$value) > 0) {
    kill = call_graph_url(me$token,
                          paste("https://graph.microsoft.com/v1.0/drives/",
                                group_drive_id,
                                "/items/",
                                excel_file_id, 
                                "/workbook/tables/",
                                tables$value[[1]]$id,
                                sep = ""),
                          http_verb=c ("DELETE"),
                          body=toJSON (list(), auto_unbox = TRUE), 
                          encode = "raw")  
  } 
  
  clear = call_graph_url(me$token,
                         paste("https://graph.microsoft.com/v1.0/drives/",
                               group_drive_id,
                               "/items/",
                               excel_file_id, 
                               "/workbook/worksheets/tasks/range(address='A1:Z1000')/clear",
                               sep = ""),
                         http_verb=c ("POST"),
                         body=toJSON (list(applyTo = "All"), auto_unbox = TRUE), 
                         encode = "raw")
  
  new_table = call_graph_url(me$token,
                             paste("https://graph.microsoft.com/v1.0/drives/",
                                   group_drive_id,
                                   "/items/",
                                   excel_file_id, 
                                   "/workbook/tables/add",
                                   sep = ""),
                             body=toJSON (list(address=paste("tasks!A1:A", length(task_tbl[[1]]), sep=""), "has-headers"=TRUE), auto_unbox = TRUE), 
                             http_verb=c ("POST"),
                             encode = "raw")
  new_table
}



add_table_column = function (column_name, new_table, file_id, group_drive_id, value_vec) {
  temp_list = sapply(as.list(value_vec), function(x) list(list(x)))
  temp_list = append(list(list(column_name)), temp_list)
  index = 0
  if(column_name == "link") {
    body_prep = toJSON (list(name = column_name, values=temp_list, index = 2), auto_unbox = TRUE)
  } else {
    body_prep = toJSON (list(name = column_name, values=temp_list), auto_unbox = TRUE)
  }
  new_column = call_graph_url(me$token,
                              paste("https://graph.microsoft.com/v1.0/drives/",
                                    group_drive_id,
                                    "/items/",
                                    file_id, 
                                    "/workbook/tables/",
                                    new_table$id, 
                                    "/columns",
                                    sep = ""),
                              body=body_prep, 
                              http_verb=c ("POST"),
                              encode = "raw")
}

update_excel = function(excel_file_id, task_tbl) {
  print("going to empty excel table")
  new_table = empty_table(excel_file_id, task_tbl)
  print("going to insert columns")
  for (numCol in seq(length(task_tbl))) {
    col_name = names(task_tbl)[[numCol]]
    print(col_name)
    col_value_vec = task_tbl[[numCol]]
    add_table_column(col_name, new_table, excel_file_id, group_drive_id, col_value_vec)
  }
}

print("going to update the active excel")
active_tasks = tasks_with_bucket_names %>%
  filter(Assigned == "unassigned" |Assigned == "AG" | Assigned == "MN" | Assigned == "DP" | Assigned == "ZV" | Assigned == "BG") %>%
  filter(Completed == "")

update_file = call_graph_url(me$token,
                             paste("https://graph.microsoft.com/v1.0/drives/",group_drive_id,"/root/search(q='{Active+TSTA+Tasks.xlsx}')", sep="")) 

file_id = update_file$value[[which(sapply(update_file$value, function(x) x$name == "Active TSTA Tasks.xlsx"))]]$id

#update_excel(file_id, active_tasks)

print("going to update the all excel")
all_tasks = tasks_with_bucket_names %>%
  filter(Assigned == "unassigned" |Assigned == "AG" | Assigned == "MN" | Assigned == "DP" | Assigned == "ZV" | Assigned == "BG")

update_file = call_graph_url(me$token,
                             paste("https://graph.microsoft.com/v1.0/drives/",group_drive_id,"/root/search(q='{All+TSTA+Tasks.xlsx}')", sep="")) 

file_id = update_file$value[[which(sapply(update_file$value, function(x) x$name == "All TSTA Tasks.xlsx"))[[1]]]]$id

update_excel(file_id, all_tasks)


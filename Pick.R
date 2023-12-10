# 須確定第四行為年齡
# 須確定第五行為排名

install.packages("openxlsx","dplyr")
library(openxlsx)
library(dplyr)

KeyWord = ".xlsx"

Number = length(list.files(pattern = KeyWord))
Result = matrix(0,Number,9)

for (ii in c(1:Number)){
  FileName = list.files(pattern = KeyWord)[ii]
  ReName = gsub(".xlsx","",FileName)
  Data = read.xlsx(FileName, sheet = 1, colNames = F)
  D0 = Data %>%
    dplyr::rename(age = X4, rank = X5) %>%
    dplyr::select(age, rank) %>%
    dplyr::filter(rank != "r") %>%
    dplyr::mutate_all(as.numeric)
  
  D150 = D0 %>%
    dplyr::filter(rank <= 150)
  
  C1 = D150$age[1] # 第一次150的年齡
  C2 = D150$age[1]-D0$age[1] # 第一次150所花時間(出道開始算)
  
  C5 = min(D0$rank)
  
  rank_now = D0$rank
  rank_last = c(10000,rank_now[-length(rank_now)])
  target_indexs = rank_now<=150&rank_last>=150
  target_index2 = rank_now<=30&rank_last>=30
  
  target_ages = D0$age[target_indexs] #符合前後rank經過150的年紀們，可能為複數個
  target_age2 = D0$age[target_index2]
  age_best_rank = D0$age[which.min(D0$rank)]
  target_ages = target_ages[target_ages<age_best_rank] #篩選最佳成績以前
  target_age2 = target_age2[target_age2<age_best_rank]
  final_answer_age = max(target_ages)
  final_answer_age2= max(target_age2)
  
  C3 = final_answer_age2-final_answer_age # 150到30所花時間
  C4 = age_best_rank-final_answer_age # 150到最高所花時間
  
  Result[ii,1] = ReName                       # 選手人名
  Result[ii,2] = round(C1,2)                  # 第一次150的年齡
  Result[ii,3] = round(C2,2)                  # 出道到第一次150所花的時間
  Result[ii,4] = round(C3,2)                  # 150到30所花時間
  Result[ii,5] = round(C4,2)                  # 150到最高所花時間
  Result[ii,6] = round(final_answer_age,2)    # 開始衝刺的時間 (爬到最高名次的那次150是幾歲)
  Result[ii,7] = round(final_answer_age2,2)   # 到達30的年齡 (爬到最高名次的那次30是幾歲)
  Result[ii,8] = round(age_best_rank,2)       # 最高名次時的年齡
  Result[ii,9] = round(C5,2)                  # 最高名次
  
  print(paste0("process = ", ii, ", total =" ,Number))
}

write.table(Result, file = "Pick.txt", col.names = F, row.names = F, quote = F)



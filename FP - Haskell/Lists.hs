import Data.List
--Самодвойственная задача
isSelfdual [] = True
isSelfdual mylist = 
    if (head mylist) /= (last mylist) then isSelfdual (init (tail mylist))
    else False

--Месяцы с праздниками
        
findMaxMonths monthList = 
    nub (map (\(d,m) -> m) (filter (\(d,month) -> maxMonths monthList == curMonth month monthList) monthList)) where
    curMonth month monthList = 
        length (filter (\(d,m) -> m==month) monthList)
    maxMonths monthList = maxMonths monthList 0 where
        maxMonths [] curMax = curMax
        maxMonths ((d,m):t) curMax =
            if (curMonth m ((d,m):t)) > curMax then maxMonths t (curMonth m ((d,m):t))
            else maxMonths t curMax

--Подстрока
isSubstr str substr = isSubstr' str str substr substr where
    isSubstr' str strLeft substr [] = True
    isSubstr' [] strLeft substr substrLeft = False
    isSubstr' str [] substr substrLeft = isSubstr' (tail str) (tail str) substr substrLeft
    isSubstr' str strLeft substr substrLeft =
        if (head strLeft) == (head substrLeft) then isSubstr' str (tail strLeft) substr (tail substrLeft)
        else isSubstr' (tail str) (tail str) substr substrLeft

--Считалочка
countFifth list = countFifth list list 1 (head list) where
    countFifth [] reducingList curNum lastName = lastName
    countFifth list [] curNum lastName = countFifth list list curNum lastName
    countFifth list reducingList curNum lastName =
        if curNum == 5 then countFifth (delete (head reducingList) list) (tail reducingList) 1 (head reducingList)
        else countFifth list (tail reducingList) (curNum+1) (head reducingList)
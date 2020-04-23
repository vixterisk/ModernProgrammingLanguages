import Data.List
import System.IO
import Data.Function
--Получает список всех вершин из исходных данных
getAllNodes list = nub $ concat $ map (\(a,b,c) ->[a, b]) list

--Инициализирует вершины метками при начале алгоритма (вершина, метка)
markNodes startNode edges =  filter (\(x, _) -> x /= startNode) (map (\x -> (x, -1)) (getAllNodes edges))

--Находит минимальную метку
minDist marks = minimum (map (\(_,mark) -> mark) marks)

--Находит вторую минимальную метку
sndMinDist marks = 
    if listWithoutMin /= [] then minimum listWithoutMin
    else -1 where 
    listWithoutMin = filter (\mark -> minDist marks /= mark) (map (\(_,mark) -> mark)  marks)

--Находит вершины с минимальной меткой
findMinEdge marks = 
    if listWithMin == [] then [(-1,-1)]
    else listWithMin where
    listWithMin = filter (\(_,mark) -> (minDist marks) == mark) marks

--Находит вершины со второй минимальной меткой
findSndMinEdge marks =
    if listWithoutMin == [] then [(-1,-1)]
    else findMinEdge listWithoutMin where
    listWithoutMin = (filter (\(_,mark) -> (minDist marks) /= mark) marks)

--Проверяет, есть ли ребро между двумя вершинами
isEdgeBetweenNodes (v1,v2,d) node1 node2 = v1 == node1 && v2 == node2

--Расстояние между текущей вершиной и выбранной
findEdge curNode chosenNode edges =
    if curNode == chosenNode then 0 else
        if resultList == [] then -1 else minimum resultList where
    resultList = map (\(_,_,d) -> d) (filter (\(node1,node2,d) -> isEdgeBetweenNodes (node1,node2,d) curNode chosenNode) edges)

--Обновляет метки соседних вершин
updateMarks curNode curDist [] edges = []
updateMarks curNode curDist ((v,mark):t) edges =
    if dist == -1 then (v, mark):(updateMarks curNode curDist t edges)
    else
        if mark == -1 then (v, dist + curDist):(updateMarks curNode curDist t edges)
        else if (dist + curDist < mark) then (v, dist + curDist):(updateMarks curNode curDist t edges)
            else (v, mark):(updateMarks curNode curDist t edges) where
    dist = findEdge curNode v edges

--Находим следующую вершину, которая станет текущей + убираем ее из списка меток
reduceMarks marks min = reduceMarks' marks min [] where
    reduceMarks' ((v,d):t) (minV,minD) acc =
        if (minD == d) then union acc t
        else reduceMarks' t (minV,minD) ((v,d):acc)

dijkstra startNode desiredNode edges = dijkstra' startNode 0 desiredNode (markNodes startNode edges) edges where
    dijkstra' curNode curDist desiredNode marks edges =
        if curNode == desiredNode then curDist
        else
            if minD /= -1 then dijkstra' minV minD desiredNode (reduceMarks newMarks (minV,minD)) edges 
            else if sndMinD /= -1 then dijkstra' sndMinV sndMinD desiredNode (reduceMarks newMarks (sndMinV,sndMinD)) edges
            else -1 where
        newMarks = updateMarks curNode curDist marks edges
        (minV,minD) = head $ findMinEdge newMarks
        (sndMinV,sndMinD) = head $ findSndMinEdge newMarks

parseToEdges[] acc = acc
parseToEdges list acc =
    parseToEdges (tail $ tail $ tail list) ((node1,node2,dist):acc) where
    node1 = read (head list)
    node2 = read (head $ tail list)
    dist = read (head $ tail $ tail list) :: Int

main = do
    filename <- getLine
    handle <- openFile filename ReadMode
    contents <- hGetContents handle
    let stringsArray = words contents
    let startNode = read (head stringsArray) :: Int
    let desiredNode = read (head $ tail stringsArray) :: Int
    let edges = parseToEdges (tail $ tail stringsArray) []
    writeFile ((take (length filename - 4) filename)++"-haskell-output.txt")  $ show $ dijkstra startNode desiredNode edges
import Data.List
import Data.Maybe

--Tree Structure
data Expr = Const Int | Add Expr Expr | Mult Expr Expr | Sub Expr Expr | Div Expr Expr deriving (Eq)

--Show
instance Show Expr where
    show (Const x) = show x
    show (Add x y) = "("++show x++" + "++show y++")"
    show (Mult x y) = "("++show x++" * "++show y++")"
    show (Sub x y) = "("++show x++" - "++show y++")"
    show (Div x y) = "("++show x++" / "++show y++")"

--Show in Reverse Polish Notation
showRPN (Add expr1 expr2) = showRPN expr1++" "++showRPN expr2++" "++"+"
showRPN (Sub expr1 expr2) = showRPN expr1++" "++showRPN expr2++" "++"-"
showRPN (Mult expr1 expr2) = showRPN expr1++" "++showRPN expr2++" "++"*"
showRPN (Div expr1 expr2) = showRPN expr1++" "++showRPN expr2++" "++"/"
showRPN (Const x) = show x

--Calculate
calcTree (Const x) = x
calcTree (Add e1 e2) = (calcTree e1) + (calcTree e2)
calcTree (Mult e1 e2) = (calcTree e1) * (calcTree e2)
calcTree (Sub e1 e2)  = (calcTree e1) - (calcTree e2)
calcTree (Div e1 e2)  = (calcTree e1) `div` (calcTree e2)
calc str = calcTree (parse str)

solveRPN = head . foldl foldingFunction [] . words  
    where   foldingFunction (x:y:ys) "*" = (x * y):ys
            foldingFunction (x:y:ys) "/" =  (x `div` y):ys
            foldingFunction (x:y:ys) "+" = (x + y):ys 
            foldingFunction (x:y:ys) "-" = (y - x):ys  
            foldingFunction xs numberString = read numberString:xs  

--Parse expression string in Tree
parse str = parse' (reverse (words str)) where
    parse' lst = 
        let firstEntry (a,b) lst = fromMaybe "" (find (\c -> c == a || c == b) lst) in
        let firstEntryIndex (a,b) lst = fromMaybe (-1) (findIndex (\c -> c == a || c == b) lst) in
        case (firstEntry ("+","-") lst) of
        "-" -> Sub (parse' (tail (drop (firstEntryIndex ("+","-") lst) lst))) (parse' (take (firstEntryIndex ("+","-") lst) lst)) 
        "+" -> Add (parse' (tail (drop (firstEntryIndex ("+","-") lst) lst))) (parse' (take (firstEntryIndex ("+","-") lst) lst))
        _ ->
            case (firstEntry ("*","/") lst) of
            "*" -> Mult (parse' (tail (drop (firstEntryIndex ("*","/") lst) lst))) (parse' (take (firstEntryIndex ("*","/") lst) lst))
            "/" -> Div (parse' (tail (drop (firstEntryIndex ("*","/") lst) lst))) (parse' (take (firstEntryIndex ("*","/") lst) lst))
            _ -> Const(read (head lst) :: Int) 
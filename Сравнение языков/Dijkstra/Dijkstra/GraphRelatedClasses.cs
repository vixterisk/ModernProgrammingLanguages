using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Dijkstra
{
    class Graph
    {
        public List<Edge> Egdes { get; set; }
        public List<string> Nodes { get; set; }
        public Graph(List<Edge> egdes)
        {
            Egdes = egdes;
            var nodeDict = new Dictionary<string, bool>();
            foreach (var edge in Egdes) 
            {
                nodeDict[edge.NodeFrom] = true;
                nodeDict[edge.NodeTo] = true;
            }
            Nodes = nodeDict.Keys.ToList();
        }
        public int Dijkstra(string startNode, string desiredNode)
        {
            //Помечаем вершины начальными метками (бесконечность как -1)
            var marks = GetMarksForDijkstra(startNode);
            var visitedNodes = new Dictionary<string, bool>();
            //На начало работы все вершины помечаются как непосещенные
            foreach (var node in Nodes)
                visitedNodes[node] = false; 
            string currentNode = startNode;
            //Продолжаем обрабатывать граф, пока текущей вершиной не станет искомая, либо пока не будут посещены все вершины
            while (desiredNode != currentNode || !AllNodesVisited(visitedNodes))
            {
                bool noReachableNodes = true;
                var min = Int32.MaxValue;
                //Выбираем текущую еще не посещенную вершину, имеющую минимальную метку и достижимую из стартовой вершины
                foreach (var node in Nodes)
                {
                    if (marks[node] != -1 && marks[node] < min && visitedNodes[node] == false)
                    { min = marks[node]; currentNode = node; noReachableNodes = false; }
                }
                //Если ни разу не заходили в предыдущее условие, закончились достижимые вершины
                if (noReachableNodes) break;
                visitedNodes[currentNode] = true;
                //Выбираем соседей текущей вершины
                var adjacentNodes = GetAdjacentNodesInOrientedGraph(currentNode);
                adjacentNodes.Sort(Edge.CompareEdgesByDistance);
                foreach (var adjacentNode in adjacentNodes)
                {
                    //Если новая длина пути, равная сумме значений текущей метки и длины ребра, меньше метки соседа, перезаписываем метку соседа
                    if (marks[adjacentNode.NodeTo] == -1 || marks[adjacentNode.NodeTo] > marks[currentNode] + adjacentNode.Distance) marks[adjacentNode.NodeTo] = marks[currentNode] + adjacentNode.Distance;
                }
            }
            //Если в графе нет ребер или заданная вершина не найдена среди вершин графа, возвращаем -1
            if (marks.Count == 0 || !marks.ContainsKey(desiredNode)) return -1;
            return marks[desiredNode];
        }
        private bool AllNodesVisited(Dictionary<string, bool> visitedNodes)
        {
            foreach (var node in Nodes)
            {
                if (!visitedNodes[node]) return false;
            }
            return true;
        }
        private Dictionary<string, int> GetMarksForDijkstra(string startNode)
        {
            var marks = new Dictionary<string, int>();
            foreach (var node in Nodes)
            {
                if (node == startNode) marks[node] = 0;
                else marks.Add(node, -1);
            }
            return marks;
        }
        List<Edge> GetAdjacentNodesInOrientedGraph(string node)
        {
            var result = new List<Edge>();
            foreach (var edge in this.Egdes)
            {
                if (edge.NodeFrom == node) result.Add(edge);
            }
            return result;
        }
    }
    class Edge
    {
        public string NodeFrom { get; set; }
        public string NodeTo { get; set; }
        public int Distance { get; set; }
        public Edge(string NodeFrom, string NodeTo, int Distance)
        {
            this.NodeFrom = NodeFrom;
            this.NodeTo = NodeTo;
            this.Distance = Distance;
        }
        public static int CompareEdgesByDistance(Edge edge1, Edge edge2)
        {
            if (edge1.Distance > edge2.Distance) return 1;
            else if (edge1.Distance == edge2.Distance) return 0;
            else return -1;
        }
    }
}

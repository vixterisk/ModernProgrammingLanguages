using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Dijkstra
{
    class Program
    {
        static void ParseToGraphAndSolveDijkstra(string filename)
        {
            filename = "C:\\Users\\vixterisk\\Desktop\\MPL\\Сравнение языков\\"+filename;
            var edges = new List<Edge>();
            var input = File.ReadAllLines(filename);
            var startNode = input[0];
            var desiredNode = input[1];
            for (int i = 2; i < input.Length; i++)
            {
                var str = input[i].Split();
                edges.Add(new Edge(str[0], str[1], Int32.Parse(str[2])));
            }
            var graph = new Graph(edges);
            var result = graph.Dijkstra(startNode, desiredNode);
            string name = filename.Substring(0, filename.Length - 4) + "-CSharp-output.txt";
            File.WriteAllLines(name, new string[] { result.ToString() });
        }
        static void GenerateReallyBigGraph()
        {
            StringBuilder str = new StringBuilder();
            for (int i = 1; i < 100; i++)
                for (int j = i + 1; j < 100; j++)
                {
                    str.Append(i.ToString() + " " + j.ToString() + " " + 10 + "\n");
                }
            File.WriteAllText("big.txt", str.ToString());
        }

        static void Main(string[] args)
        {
           //GenerateReallyBigGraph();
           Console.WriteLine("Write input file name: \n");
           var filename = Console.ReadLine();
           ParseToGraphAndSolveDijkstra(filename);
        }
    }
}

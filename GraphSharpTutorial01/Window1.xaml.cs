using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using GraphSharp;
using QuickGraph;
using GraphSharp.Controls;
using GraphSharpTutorial01;
using GraphSharp.Algorithms.Layout.Simple.Tree;
using System.Threading;
using WPFExtensions.Controls;

namespace WpfApplication1
{

    public partial class MainWindow : Window
    {
        public IBidirectionalGraph<object, QuickGraph.IEdge<object>> Graph { get; set; }


        public MainWindow()
        {
            this.InitializeComponent();
            InitializeComponent();
            
        }

        

        private void Button_Click(object sender, RoutedEventArgs e)
        {

            Global.NodeName = NodeNameTextBox.Text;
            var g = new BidirectionalGraph<object, QuickGraph.IEdge<object>>();

            var re = new BidirectionalGraph<string, TaggedEdge<string, string>>();

            FromExcel FE = new FromExcel();

            if (VariantComboBox.Text=="Потомки")
            {
                FE.InputParam(1);
            }
            if (VariantComboBox.Text == "Родители")
            {
                FE.InputParam(2);
            }


            IList<Object> vertices = new List<Object>();
            for (int i = 0; i < Global.Graph.Count; i++)
            {
                vertices.Add(Global.Graph[i].Parent);
                vertices.Add(Global.Graph[i].Name);
            }

            for (int i = 0; i < vertices.Count; i++)
            {                
                g.AddVerticesAndEdge(new MyEdge(vertices[i+1], vertices[i])
                {
                    Id = i.ToString(),
                });

                i++;
            }

            
            Graph = g;

            graphLayout.Graph = Graph;
            //Zoom.Mode = 1;
            DataContext = graphLayout.Graph;
            AddEventsToGraph();            
        }

        private void AddEventsToGraph()
        {   

            foreach (var v in this.graphLayout.Children)
            {           
                if (v is EdgeControl)
                {
                    EdgeControl vc = (EdgeControl)v;
                    Source=(string)vc.Source.Vertex;
                    Target = (string)vc.Target.Vertex;
                    vc.MouseEnter += MainWindow_MouseEnter;
                    vc.MouseLeave += MainWindow_MouseLeave;

                }
            }
        }

        GraphNode node = new GraphNode();
        void MainWindow_MouseEnter(object sender, MouseEventArgs e)
        {           
            node.ShowToolTip(Source, Target);
        }
        void MainWindow_MouseLeave(object sender, MouseEventArgs e)
        {
            
            node.CloseToolTip();
        }

        public string Source { get; set; }
        public string Target { get; set; }

    }
        

    public class GraphNode
    {
        private ToolTip tp;
        
        public GraphNode()
        {            
            tp = new ToolTip();
        }
        public void ShowToolTip( string str1, string str2 )
        {
            for (int i = 0; i < Global.Graph.Count;i++ )
            {
                if (Global.Graph[i].Name==str1 && Global.Graph[i].Parent==str2)
                {
                    tp.Content = (string)Global.Graph[i].Value;
                    tp.IsOpen = true;
                }
            }
                
            
        }
        public void CloseToolTip()
        {
            this.tp.IsOpen = false;         
        }
    }

    public class MyEdge : TypedEdge<Object>
    {
        public String Id { get; set; }

        public MyEdge(Object source, Object target) : base(source, target, EdgeTypes.General) { }
    }

    
    public class EdgeColorConverter : IValueConverter
    {

        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            return new SolidColorBrush((Color)value);
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}

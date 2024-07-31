using TabularEditor.BestPracticeAnalyzer; 
//using TabularEditor.Shared.BPA; // Use this instead of the first line when using Tabular Editor 3

var bpa = new Analyzer();
bpa.SetModel(Model);

var sb = new System.Text.StringBuilder();
string newline = Environment.NewLine;

sb.Append("RuleCategory" + '\t' + "RuleName" + '\t' + "ObjectName" + '\t' + "ObjectType" + '\t' + "RuleSeverity" + '\t' + "HasFixExpression" + newline);

foreach (var a in bpa.AnalyzeAll().ToList())
{
    sb.Append(a.Rule.Category + '\t' + a.RuleName + '\t' + a.ObjectName + '\t' + a.ObjectType + '\t' + a.Rule.Severity + '\t' + a.CanFix + newline);
}

sb.Output();

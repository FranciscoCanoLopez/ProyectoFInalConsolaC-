
Namespace ProyectoFInalConsolaC_
    Friend Class CsvRow
                ''' Cannot convert PropertyDeclarationSyntax, System.InvalidCastException: No se puede convertir un objeto de tipo 'Microsoft.CodeAnalysis.VisualBasic.Syntax.EmptyStatementSyntax' al tipo 'Microsoft.CodeAnalysis.VisualBasic.Syntax.ExpressionSyntax'.
'''    en ICSharpCode.CodeConverter.VB.NodesVisitor.VisitPropertyDeclaration(PropertyDeclarationSyntax node)
'''    en Microsoft.CodeAnalysis.CSharp.CSharpSyntaxVisitor`1.Visit(SyntaxNode node)
'''    en ICSharpCode.CodeConverter.VB.CommentConvertingVisitorWrapper`1.Accept(SyntaxNode csNode, Boolean addSourceMapping)
''' 
''' Input:
'''         public System.Collections.Generic.Dictionary<string, string> Fields { get; set; } = new();
''' 
''' 
        Public Property Player As String?
        Public Property Nationality As String?
        Public Property Position As String?
        Public Property Club As String?
        Public Property Age As String?
        Public Property Matches As String?
        Public Property Goals As String?
    End Class
End Namespace

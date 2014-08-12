﻿'------------------------------------------------------------------------------
' <autogenerated>
'     This code was generated by a tool.
'     Runtime Version: 1.1.4322.573
'
'     Changes to this file may cause incorrect behavior and will be lost if 
'     the code is regenerated.
' </autogenerated>
'------------------------------------------------------------------------------

Option Strict Off
Option Explicit On

Imports System
Imports System.ComponentModel
Imports System.Diagnostics
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.Xml.Serialization

'
'This source code was auto-generated by Microsoft.VSDesigner, Version 1.1.4322.573.
'
Namespace net.webservicex.www
    
    '<remarks/>
    <System.Diagnostics.DebuggerStepThroughAttribute(),  _
     System.ComponentModel.DesignerCategoryAttribute("code"),  _
     System.Web.Services.WebServiceBindingAttribute(Name:="TranslationServiceSoap", [Namespace]:="http://www.webservicex.net/")>  _
    Public Class TranslationService
        Inherits System.Web.Services.Protocols.SoapHttpClientProtocol
        
        '<remarks/>
        Public Sub New()
            MyBase.New
            Me.Url = "http://www.webservicex.net/TranslateService.asmx"
        End Sub
        
        '<remarks/>
        <System.Web.Services.Protocols.SoapDocumentMethodAttribute("http://www.webservicex.net/Translate", RequestNamespace:="http://www.webservicex.net/", ResponseNamespace:="http://www.webservicex.net/", Use:=System.Web.Services.Description.SoapBindingUse.Literal, ParameterStyle:=System.Web.Services.Protocols.SoapParameterStyle.Wrapped)>  _
        Public Function Translate(ByVal LanguageMode As Language, ByVal [Text] As String) As String
            Dim results() As Object = Me.Invoke("Translate", New Object() {LanguageMode, [Text]})
            Return CType(results(0),String)
        End Function
        
        '<remarks/>
        Public Function BeginTranslate(ByVal LanguageMode As Language, ByVal [Text] As String, ByVal callback As System.AsyncCallback, ByVal asyncState As Object) As System.IAsyncResult
            Return Me.BeginInvoke("Translate", New Object() {LanguageMode, [Text]}, callback, asyncState)
        End Function
        
        '<remarks/>
        Public Function EndTranslate(ByVal asyncResult As System.IAsyncResult) As String
            Dim results() As Object = Me.EndInvoke(asyncResult)
            Return CType(results(0),String)
        End Function
    End Class
    
    '<remarks/>
    <System.Xml.Serialization.XmlTypeAttribute([Namespace]:="http://www.webservicex.net/")>  _
    Public Enum Language
        
        '<remarks/>
        EnglishTOChinese
        
        '<remarks/>
        EnglishTOFrench
        
        '<remarks/>
        EnglishTOGerman
        
        '<remarks/>
        EnglishTOItalian
        
        '<remarks/>
        EnglishTOJapanese
        
        '<remarks/>
        EnglishTOKorean
        
        '<remarks/>
        EnglishTOPortuguese
        
        '<remarks/>
        EnglishTOSpanish
        
        '<remarks/>
        ChineseTOEnglish
        
        '<remarks/>
        FrenchTOEnglish
        
        '<remarks/>
        FrenchTOGerman
        
        '<remarks/>
        GermanTOEnglish
        
        '<remarks/>
        GermanTOFrench
        
        '<remarks/>
        ItalianTOEnglish
        
        '<remarks/>
        JapaneseTOEnglish
        
        '<remarks/>
        KoreanTOEnglish
        
        '<remarks/>
        PortugueseTOEnglish
        
        '<remarks/>
        RussianTOEnglish
        
        '<remarks/>
        SpanishTOEnglish
    End Enum
End Namespace
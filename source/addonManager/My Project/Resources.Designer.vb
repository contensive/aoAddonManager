﻿'------------------------------------------------------------------------------
' <auto-generated>
'     This code was generated by a tool.
'     Runtime Version:4.0.30319.42000
'
'     Changes to this file may cause incorrect behavior and will be lost if
'     the code is regenerated.
' </auto-generated>
'------------------------------------------------------------------------------

Option Strict On
Option Explicit On

Imports System

Namespace My.Resources
    
    'This class was auto-generated by the StronglyTypedResourceBuilder
    'class via a tool like ResGen or Visual Studio.
    'To add or remove a member, edit your .ResX file then rerun ResGen
    'with the /str option, or rebuild your VS project.
    '''<summary>
    '''  A strongly-typed resource class, for looking up localized strings, etc.
    '''</summary>
    <Global.System.CodeDom.Compiler.GeneratedCodeAttribute("System.Resources.Tools.StronglyTypedResourceBuilder", "15.0.0.0"),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Runtime.CompilerServices.CompilerGeneratedAttribute(),  _
     Global.Microsoft.VisualBasic.HideModuleNameAttribute()>  _
    Friend Module Resources
        
        Private resourceMan As Global.System.Resources.ResourceManager
        
        Private resourceCulture As Global.System.Globalization.CultureInfo
        
        '''<summary>
        '''  Returns the cached ResourceManager instance used by this class.
        '''</summary>
        <Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)>  _
        Friend ReadOnly Property ResourceManager() As Global.System.Resources.ResourceManager
            Get
                If Object.ReferenceEquals(resourceMan, Nothing) Then
                    Dim temp As Global.System.Resources.ResourceManager = New Global.System.Resources.ResourceManager("Resources", GetType(Resources).Assembly)
                    resourceMan = temp
                End If
                Return resourceMan
            End Get
        End Property
        
        '''<summary>
        '''  Overrides the current thread's CurrentUICulture property for all
        '''  resource lookups using this strongly typed resource class.
        '''</summary>
        <Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)>  _
        Friend Property Culture() As Global.System.Globalization.CultureInfo
            Get
                Return resourceCulture
            End Get
            Set
                resourceCulture = value
            End Set
        End Property
        
        '''<summary>
        '''  Looks up a localized string similar to &lt;div class=&quot;amLibraryListCell&quot;&gt;
        '''			&lt;table&gt;
        '''				&lt;tbody&gt;
        '''					&lt;tr&gt;
        '''						&lt;td class=&quot;amImage&quot;&gt;&lt;img src=&quot;##imageLink##&quot; data-image=&quot;la3p7077q2lx&quot;&gt;&lt;/td&gt;
        '''						&lt;td class=&quot;amBody&quot;&gt;
        '''							&lt;div class=&quot;amHeadline&quot;&gt;##name##&lt;/div&gt;
        '''							&lt;div class=&quot;amButton&quot;&gt;##button##&lt;/div&gt;
        '''							&lt;div class=&quot;amCheckbox&quot;&gt;##checkbox##&lt;/div&gt;
        '''							&lt;div class=&quot;amDate&quot;&gt;Last updated: ##date##&lt;/div&gt;
        '''							&lt;div class=&quot;amDescription&quot;&gt;##description##&lt;/div&gt;
        '''						&lt;/td&gt;
        '''					&lt;/tr&gt;
        '''				&lt;/tbody&gt;
        '''			&lt;/table&gt;
        '''&lt;/div&gt;.
        '''</summary>
        Friend ReadOnly Property AddonManagerLibraryListCell() As String
            Get
                Return ResourceManager.GetString("AddonManagerLibraryListCell", resourceCulture)
            End Get
        End Property
    End Module
End Namespace

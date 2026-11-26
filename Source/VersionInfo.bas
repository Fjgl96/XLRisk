Attribute VB_Name = "VersionInfo"
Option Explicit

' Version of this XLRisk build (Semantic Versioning, https://semver.org)
Public Const XLRISK_VERSION As String = "1.2.0-beta.2"

Public Function XLRiskVersion() As String
Attribute XLRiskVersion.VB_Description = "Returns the version of the XLRisk add-in"
    XLRiskVersion = XLRISK_VERSION
End Function

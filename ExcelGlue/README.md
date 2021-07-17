# How to create a new ExcelGlue VS project
## Visual Studio
Create a library (.NET Framework) F# project.
Select .NET Framework 4.8.

## Nugget package
Given a project name, MyProject:
1. Create a MyProject-Addin.dna file in the same directory as the .fsproj file, and the following code:
```xml
<?xml version="1.0" encoding="utf-8"?>
<DnaLibrary Name="MyProject Add-In" RuntimeVersion="v4.0" xmlns="http://schemas.excel-dna.net/addin/2018/05/dnalibrary">
      <ExternalLibrary Path="MyProject.dll" ExplicitExports="false" LoadFromBytes="true" Pack="true" IncludePdb="false" ></ExternalLibrary>
</DnaLibrary>
```
2. Add this code to the .fsproj file:
```xml
  <PropertyGroup>
        <ExcelDnaAllowPackageReferenceProjectStyle>true</ExcelDnaAllowPackageReferenceProjectStyle>
        <RunExcelDnaSetDebuggerOptions>false</RunExcelDnaSetDebuggerOptions>
        <!-- etc -->
  </PropertyGroup>
```
3. Add this code to the .fsproj file:
```xml
  <ItemGroup>
      <None Include="MyProject-Addin.dna" />  <!-- THIS -->
      <Compile Include="Library1.fs" />
      <!-- etc -->
  </ItemGroup>
```
4. Install Excel-DNA Nugget packages:  
`ExcelDna.Addin version 1.1.1`  
`ExcelDna.Integration version 1.1.0`  

## Referencing ExcelGlue dll
Add Project Reference > Browse > select `ExcelGlue.dll`.
Add `open ExcelGlue` to MyProject.fs.

## Misc
Typical ExcelDna files location:  
`C:\Users\[Username]\.nuget\packages\exceldna.addin\1.1.1\readme.txt`  
`C:\Users\[Username]\.nuget\packages\exceldna.addin\1.1.1\build\ExcelDna.AddIn.targets`  
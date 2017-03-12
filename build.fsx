//include FAKE
#r "packages/Fake/tools/FakeLib.dll"

open Fake
open Fake.AssemblyInfoFile
open Fake.Git
open Fake.GitVersionHelper
open Fake.Testing.XUnit2
open System;

//Properties
let buildDir = "./build"
let artifactsDir = getBuildParamOrDefault "artifactsDir" "./artifacts"

//AssemblyInfo
let projectName = "SimpleExcelExporter";
let authors = ["Adam Wright"]
let projectDecription = "Extract data into an excel file with ease. ";
let projectSummary = "Extract data into an excel file with ease. "
let tags = "C# Excel"

let copyright = DateTime.Now.Year.ToString()


//Version Info
let gitversionPath = @"packages\GitVersion.CommandLine\tools\GitVersion.exe"
let mutable assemblyVersion = ""
let mutable nugetVersion = ""
let mutable informationalVersion = ""
let mutable commitHash = ""
let mutable majorMinorVersion = ""

//Targets
Target "Clean" (fun _ -> 
    CleanDirs [buildDir;artifactsDir;]
)

Target "VersionLibraries" (fun _ ->
    
    let result = GitVersion (fun p -> { p with ToolPath = gitversionPath })

    assemblyVersion <- result.AssemblySemVer
    nugetVersion <- result.NuGetVersion
    informationalVersion <- result.InformationalVersion
    majorMinorVersion <- result.MajorMinorPatch + ".0"   
)

Target "BuildApp" (fun _ -> 
     
     CreateCSharpAssemblyInfo "./source/Properties/AssemblyInfo.cs"
        [Attribute.Title projectName
         Attribute.Description projectDecription
         Attribute.Product projectName
         Attribute.Copyright copyright
         Attribute.Version assemblyVersion
         Attribute.FileVersion majorMinorVersion
         Attribute.InformationalVersion informationalVersion
         Attribute.Metadata("githash", commitHash)]

     !! "source/**/*.csproj"
        |> MSBuildRelease buildDir "Build"
        |> Log "Build-Output: "
)


Target "RunTests" (fun _ -> 
    
    let xUnitExe = @"packages\xunit.runner.console\tools\xunit.console.exe"

    let xUnitHtmlOutput = (artifactsDir @@ "tests.html")

    !! "tests/**/*.csproj"
        |> MSBuildRelease buildDir "Build"
        |> Log "BuildTests-Output: "

    !! (buildDir @@ "*.Tests.dll")
    |> xUnit2 (fun p -> { p with 
                           HtmlOutputPath = Some xUnitHtmlOutput
        })
) 

Target "Package" (fun _ -> 
    
    let files = [(buildDir @@ projectName + ".dll", Some "lib/net452", None)
                 (buildDir @@ projectName + ".pdb", Some "lib/net452" , None)]

    let dependencies = [("EPPlus","4.1.0")]
    
    let nuspec =  projectName + ".nuspec"

    NuGet (fun p -> 
        {p with
              Authors = authors
              Project = projectName
              Description = projectDecription
              Summary = projectSummary
              Version = nugetVersion
              Tags = tags
              Copyright = copyright
              OutputPath = artifactsDir
              Files = files
              Dependencies = dependencies
              WorkingDir = "."
         }) 
        "SimpleExcelExporter.nuspec"
)

Target "Publish" (fun _ ->
    trace "publishing"
)


"Clean"
    ==> "VersionLibraries"
    ==> "BuildApp"
    ==> "RunTests"
    ==> "Package"
    ==> "Publish"

RunTargetOrDefault "Package"
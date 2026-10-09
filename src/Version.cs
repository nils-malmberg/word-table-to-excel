using System.Reflection;
using System.Runtime.Versioning;

// Numéro de version de l'application et du complément : c'est le SEUL endroit où il figure.
// Compilé par les deux projets (SDK .NET) et par build.cmd (compilateur C# de Windows).
[assembly: AssemblyVersion("2.1.1.0")]
[assembly: AssemblyFileVersion("2.1.1.0")]
[assembly: AssemblyInformationalVersion("2.1.1")]

// Plate-forme visée : .NET Framework 4 (comportement identique quel que soit le compilateur).
[assembly: TargetFramework(".NETFramework,Version=v4.0", FrameworkDisplayName = ".NET Framework 4")]

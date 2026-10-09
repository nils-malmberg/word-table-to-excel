using System.Reflection;
using System.Runtime.InteropServices;

// Version : src/Version.cs (commun à l'application et au complément).
[assembly: AssemblyTitle("WordTableToExcel")]
[assembly: AssemblyProduct("Tableaux Word vers Excel")]
[assembly: AssemblyCompany("WordTableToExcel")]
[assembly: AssemblyDescription("Complément COM Word : exporte les tableaux du document vers Excel et importe des tableaux Excel dans Word, en conservant la mise en forme.")]

// Seules les classes explicitement marquées [ComVisible(true)] sont exposées à COM.
[assembly: ComVisible(false)]
[assembly: Guid("5b0f3c7e-2d5a-4f0e-9a51-8f3f2f7d9c11")]

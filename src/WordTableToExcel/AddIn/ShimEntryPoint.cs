using System;
using System.Globalization;
using System.Runtime.InteropServices;
using WordTableToExcel.Infrastructure;

namespace WordTableToExcel.AddIn
{
    /// <summary>
    /// Point d'entrée appelé par le chargeur natif (WordTableToExcel.Shim32.dll / Shim64.dll) via
    /// ICLRRuntimeHost::ExecuteInDefaultAppDomain, qui impose la signature « static int Méthode(string) ».
    /// Le chargeur passe l'adresse (hexadécimale) d'une variable IUnknown* : on y écrit le pointeur
    /// IUnknown de l'objet du complément, que Word utilise ensuite comme n'importe quel objet COM.
    /// </summary>
    public static class ShimEntryPoint
    {
        public static int CreateAddIn(string resultAddress)
        {
            try
            {
                long address = long.Parse(resultAddress, NumberStyles.HexNumber, CultureInfo.InvariantCulture);
                IntPtr unknown = Marshal.GetIUnknownForObject(new Connect()); // référence déjà comptée
                Marshal.WriteIntPtr(new IntPtr(address), unknown);
                return 0;
            }
            catch (Exception ex)
            {
                Log.Error("Création du complément par le chargeur natif", ex);
                int hr = Marshal.GetHRForException(ex);
                return hr != 0 ? hr : -1;
            }
        }
    }
}

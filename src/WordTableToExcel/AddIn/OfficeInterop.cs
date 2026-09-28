using System;
using System.Runtime.InteropServices;

// Déclarations des interfaces COM d'Office utilisées par le complément.
// Elles sont redéclarées ici (mêmes GUID que les assemblys PIA « Extensibility » et « Office »)
// afin de ne dépendre d'aucune version d'Office : le complément fonctionne avec toutes.
namespace WordTableToExcel.AddIn.Interop
{
    [ComVisible(true)]
    public enum ext_ConnectMode
    {
        ext_cm_AfterStartup = 0,
        ext_cm_Startup = 1,
        ext_cm_External = 2,
        ext_cm_CommandLine = 3,
        ext_cm_Solution = 4,
        ext_cm_UISetup = 5
    }

    [ComVisible(true)]
    public enum ext_DisconnectMode
    {
        ext_dm_HostShutdown = 0,
        ext_dm_UserClosed = 1,
        ext_dm_UISetupComplete = 2,
        ext_dm_SolutionClosed = 3
    }

    /// <summary>Interface de tout complément COM Office (Office 2000 et suivants).</summary>
    [ComImport, ComVisible(true)]
    [Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744")]
    [TypeLibType((short)0x1040)]
    public interface IDTExtensibility2
    {
        [DispId(1)]
        void OnConnection([In, MarshalAs(UnmanagedType.IDispatch)] object Application, [In] ext_ConnectMode ConnectMode,
            [In, MarshalAs(UnmanagedType.IDispatch)] object AddInInst, [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

        [DispId(2)]
        void OnDisconnection([In] ext_DisconnectMode RemoveMode, [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

        [DispId(3)]
        void OnAddInsUpdate([In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

        [DispId(4)]
        void OnStartupComplete([In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

        [DispId(5)]
        void OnBeginShutdown([In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
    }

    /// <summary>Personnalisation du ruban (Office 2007 et suivants).</summary>
    [ComImport, ComVisible(true)]
    [Guid("000C0396-0000-0000-C000-000000000046")]
    [TypeLibType((short)0x1040)]
    public interface IRibbonExtensibility
    {
        [DispId(1)]
        [return: MarshalAs(UnmanagedType.BStr)]
        string GetCustomUI([In, MarshalAs(UnmanagedType.BStr)] string RibbonID);
    }

    /// <summary>Contrôle du ruban transmis aux fonctions de rappel.</summary>
    [ComImport, ComVisible(true)]
    [Guid("000C0395-0000-0000-C000-000000000046")]
    [TypeLibType((short)0x1040)]
    public interface IRibbonControl
    {
        [DispId(1)]
        string Id { [return: MarshalAs(UnmanagedType.BStr)] get; }

        [DispId(2)]
        object Context { [return: MarshalAs(UnmanagedType.IDispatch)] get; }

        [DispId(3)]
        string Tag { [return: MarshalAs(UnmanagedType.BStr)] get; }
    }

    /// <summary>Événements d'un bouton de barre d'outils (Word 2000 à 2003).</summary>
    [ComImport, ComVisible(true)]
    [Guid("000C0351-0000-0000-C000-000000000046")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface CommandBarButtonEvents
    {
        [DispId(1)]
        void Click([In, MarshalAs(UnmanagedType.IDispatch)] object Ctrl, [In, Out] ref bool CancelDefault);
    }
}

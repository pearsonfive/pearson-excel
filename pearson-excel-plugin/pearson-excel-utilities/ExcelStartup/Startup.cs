using Microsoft.Win32;
using NetOffice.ExcelApi;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using Pearson.Excel.Utilities.Extensions;
using Action = System.Action;

namespace Pearson.Excel.Plugin.ExcelStartup
{
    public class Startup
    {
        public static List<int> HwndList = new List<int>();

        public static Application StartExcel(ExcelStartupDetails details)
        {
            Application excel;

            bool excelSuccessfullyStarted;
            var attempt = 0;
            var successfullyRegisteredXlls = false;
            bool retry;

            do
            {
                attempt++;

                enableAllDisabledItems();

                excel = new Application { Visible = true };

                var hwnd = xltry(() => excel.Hwnd);
                xltry(() => excel.Workbooks.Add());
                HwndList.Add(hwnd);

                var tokenSource = new CancellationTokenSource();
                var cancellationToken = tokenSource.Token;

                var allAddinPaths = details.Addins.StandardAddins.Select(a => a.Path);

                var excelStartTask = Task.Run(() =>
                {
                    try
                    {
                        registerXlls(excel, allAddinPaths);
                        successfullyRegisteredXlls = true;
                    }
                    catch (Exception e)
                    {
                        successfullyRegisteredXlls = true;
                    }
                }, cancellationToken);

                // timed out?
                var timedOut = false;
                if (!excelStartTask.Wait(TimeSpan.FromMilliseconds(details.MaximumTimeAllowed)))
                {
                    // excel failed to start successfully in the maximum time allowed
                    tokenSource.Cancel();
                    tokenSource.Dispose();
                    timedOut = true;
                }

                // everything good or not?
                if (timedOut || !successfullyRegisteredXlls)
                {
                    excelSuccessfullyStarted = false;
                    retry = attempt < details.MaximumAttempts;
                    KillProcessByMainWindowHwnd(hwnd);
                }
                else
                {
                    excelSuccessfullyStarted = true;
                    retry = false;
                }

            } while (retry);

            // all good
            if (excelSuccessfullyStarted) return excel;

            // otherwise
            return null;
        }

        private static void registerXlls(Application excel, IEnumerable<string> paths)
        {
            try
            {
                paths.ForEach(path =>
                {
                    if (xltry(() => excel.RegisterXLL(path)))
                    {
                    }
                    else
                    {
                        throw new Exception($"Failed to register XLL [{path}]");
                    }
                });
            }
            catch (Exception e)
            {
                throw new Exception("Issue with registerXlls");
            }
        }

        private static void enableAllDisabledItems()
        {
            try
            {
                const string REGISTRY_PATH = @"Software/Microsoft/Office/14.0/Excel/Resiliency";
                var regKey = Registry.CurrentUser.CreateSubKey(REGISTRY_PATH);
                regKey?.DeleteSubKey("DisabledItems", false);
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
                throw;
            }
        }

        private const int TRY_THIS_MANY_TIMES_UNTIL_SUCCESS = 100;
        public static void xltry(Action action)
        {
            for (var i = 0; i < TRY_THIS_MANY_TIMES_UNTIL_SUCCESS; i++)
            {
                try
                {
                    action();
                    return;
                }
                catch (Exception e)
                {
                    if (!shouldRetry(e)) throw;
                    Thread.Sleep(500);
                }
            }

            throw new Exception("Excel was busy for too long.");
        }

        public static T xltry<T>(Func<T> func)
        {
            for (var i = 0; i < TRY_THIS_MANY_TIMES_UNTIL_SUCCESS; i++)
            {
                try
                {
                    return func();
                }
                catch (Exception e)
                {
                    if (!shouldRetry(e)) throw;
                    Thread.Sleep(500);
                }
            }
            throw new Exception("Excel was busy for too long.");
        }

        private const uint RPC_E_SERVERCALL_RETRYLATER = 0x8001010A;
        private const uint RPC_E_CALL_REJECTED = 0x80010001;
        private const uint VBA_E_IGNORE = 0x800AC472;

        private static bool shouldRetry(Exception e)
        {
            var exception = e;
            while (exception?.InnerException != null)
            {
                exception = exception.InnerException;
            }

            var comException = exception as COMException;
            if (comException == null) return false;

            var errorCode = (uint)comException.HResult;
            switch (errorCode)
            {
                case RPC_E_SERVERCALL_RETRYLATER:
                case VBA_E_IGNORE:
                case RPC_E_CALL_REJECTED:
                    return true;
                default:
                    return false;
            }
        }

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);
        public static void KillProcessByMainWindowHwnd(int hwnd)
        {
            GetWindowThreadProcessId((IntPtr)hwnd, out var processId);
            if (processId == 0) 
                throw new ArgumentException("Process has not been found by the given hWnd", nameof(hwnd));
            Process.GetProcessById((int)processId).Kill();
        }
    }
}
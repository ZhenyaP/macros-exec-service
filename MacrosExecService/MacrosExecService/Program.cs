using Topshelf;

namespace MacrosExecService
{
    class Program
    {
        static void Main(string[] args)
        {
            StartTopshelf();
        }

        static void StartTopshelf()
        {
            HostFactory.Run(x =>
            {
                x.Service<WebServer>(s =>
                {
                    s.ConstructUsing(name => new WebServer());
                    s.WhenStarted(tc => tc.Start());
                    s.WhenStopped(tc => tc.Stop());
                });
                x.RunAsLocalSystem();

                x.SetDescription("Web API Service to create and execute macros in Excel file.");
                x.SetDisplayName("Macro Executor Service");
                x.SetServiceName("MacroExecutor");
            });
        }
    }
}

using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(ExportToExcel.Startup))]
namespace ExportToExcel
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}

using Syncfusion.Licensing;
using System.Web;
using System.Web.Http;
using System.Web.Mvc;
using System.Web.Optimization;
using System.Web.Routing;

namespace WebApplication
{
    public class MvcApplication : HttpApplication
    {
        protected void Application_Start()
        {
            SyncfusionLicenseProvider
                .RegisterLicense("MTgzMTUwQDMxMzcyZTM0MmUzMFVoYmZhbitYWWpjclhtcTBVeUxHSzc4SkFaeW5reXVva25lc0I0VHJCQTg9");
            AreaRegistration.RegisterAllAreas();
            FilterConfig.RegisterGlobalFilters(GlobalFilters.Filters);
            RouteConfig.RegisterRoutes(RouteTable.Routes);
            BundleConfig.RegisterBundles(BundleTable.Bundles);
        }
    }
}
using Core;
using Core.Services.Class;
using Core.Services.Student;
using System.Web.Mvc;
using Unity;
using Unity.Mvc5;

namespace Demo.UI.App_Start
{
    public static partial class UnityConfig
    {
        public static void RegisterComponents()
        {
            var container = new UnityContainer();
            container.RegisterType(typeof(IGenericRepository<>), typeof(GenericRepository<>));
            container.RegisterType<IClassService, ClassService>();
            container.RegisterType<IStudentService, StudentService>();
            DependencyResolver.SetResolver(new UnityDependencyResolver(container));
        }
    }
}
using System.Collections.Generic;

namespace Core.Services.Class
{
    using Core.Models;
    public interface IClassService
    {
        Class GetById(long? id);
        IEnumerable<Class> GetAll(int page, int pageSize, ref int count);
        long Add(Class entity);
        bool Update(Class entity);
        bool Delete(Class entity);
    }
}

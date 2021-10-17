using System;
using System.Collections.Generic;


namespace Core.Services.Class
{
    using Core.Models;
    public class ClassService : IClassService
    {
        private readonly IGenericRepository<Class> _classRepo;

        public ClassService(IGenericRepository<Class> classRepo)
        {
            _classRepo = classRepo;
        }
        public long Add(Class entity)
        {
            if (entity == null)
                return 0;
            try
            {
                _classRepo.Add(entity);
                return entity.Id;
            }
            catch (Exception e)
            {
                return 0;
            }
        }

        public bool Delete(Class entity)
        {
            throw new NotImplementedException();
        }

        public IEnumerable<Class> GetAll(int page, int pageSize, ref int count)
        {
            throw new NotImplementedException();
        }

        public Class GetById(long? id)
        {
            throw new NotImplementedException();
        }

        public bool Update(Class entity)
        {
            throw new NotImplementedException();
        }
    }
}

using System.Collections.Generic;
using System.Data.Entity;
using System.Linq;
using System.Threading.Tasks;

namespace Core
{
    public interface IGenericRepository<T> where T : class
    {
        Task<T> GetById(long? id);
        IEnumerable<T> GetAll();
        Task Add(T entity);
        Task Update(T entity);
        Task Delete(T entity);
    }

    public class GenericRepository<T> : IGenericRepository<T> where T : class
    {
        //your context here ^^
        //private readonly MyEntities _entities;
        protected DbSet<T> _dbset;

        public GenericRepository()
        {
            //_entities = entities;
           // _dbset = entities.Set<T>();
        }

        public async Task Add(T entity)
        {
            //_dbset.Add(entity);
            //await _entities.SaveChangesAsync();
        }

        public async Task Delete(T entity)
        {
            //_dbset.Remove(entity);
            //await _entities.SaveChangesAsync();
        }



        public async Task Update(T entity)
        {
            //_entities.Entry(entity).State = EntityState.Modified;
            //await _entities.SaveChangesAsync();

        }

        public async Task<T> GetById(long? id)
        {
            try
            {
                return _dbset.Find(id);
            }
            catch
            {
                return null;
            }
        }

        public IEnumerable<T> GetAll()
        {
            return _dbset.AsEnumerable();
        }
    }
}

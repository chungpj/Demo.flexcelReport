using System;
using System.Collections.Generic;

namespace Core.Services.Student
{
    using Core.Models;
    public class StudentService : IStudentService
    {
        private readonly IGenericRepository<Student> _studentRepo;
        public StudentService(IGenericRepository<Student> studentRepo)
        {
            _studentRepo = studentRepo;
        }
        public long Add(Student entity)
        {
            throw new NotImplementedException();
        }

        public bool Delete(Student entity)
        {
            throw new NotImplementedException();
        }

        public IEnumerable<Student> GetAll(int page, int pageSize, ref int count)
        {
            throw new NotImplementedException();
        }

        public IEnumerable<Student> GetAll()
        {
            var classModels = this.InitClassData();
            return InitStudentData(classModels);
        }

        #region Fake Data
        private List<Class> InitClassData()
        {
            return new List<Class>
            {
                new Class {Id=11, Name = "5A"},
                new Class {Id=12, Name = "5B"},
                new Class {Id=13, Name = "5C"},
                new Class {Id=14, Name = "5D"},
            };
        }

        private List<Student> InitStudentData(List<Class> classModels)
        {
            var rs = new List<Student>();
            var rand = new Random();
            foreach (var classModel in classModels)
            {
                for (int i = 1; i <= rand.Next(3, 10); i++)
                {
                    rs.Add(new Student
                    {
                        Id = i,
                        Name = "Student " + i,
                        Math = rand.Next(0, 10),
                        Literature = rand.Next(0, 10),
                        Class = classModel
                    });
                }
            }
            return rs;
        }
        #endregion

        public Student GetById(long? id)
        {
            throw new NotImplementedException();
        }

        public bool Update(Student entity)
        {
            throw new NotImplementedException();
        }
    }
}

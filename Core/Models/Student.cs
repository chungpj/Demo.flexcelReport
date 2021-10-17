namespace Core.Models
{
    public class Student
    {
        public int Id { get; set; }
        public string Code { get; set; }
        public string Name { get; set; }
        public string Address { get; set; }
        public string PhoneNumber { get; set; }
        public int? Math { get; set; }
        public int? Literature { get; set; }
        public int Class_Id { get; set; }

        public int Average
        {
            get
            {
                var math = Math ?? 0;
                var literature = Literature ?? 0;
                return (math + literature) / 2;
            }
        }
        private Class _class;

        public Class Class
        {
            get { return _class; }
            set { 
                _class = value;
                Class_Id = _class.Id;
            }
        }
    }
}

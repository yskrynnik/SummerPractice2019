using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;


namespace SummerPractice
{
    [DataContract]
    class StudMark
    {
        // Поля класса необходимые для десериализации данных из файла
        [DataMember]
        public int id;
        [DataMember]
        public string last_name_ukr;
        [DataMember]
        public string name_ukr;
        [DataMember]
        public string group_number;
        [DataMember]
        public string short_name;
        [DataMember]
        public string name;
        [DataMember]
        public string check_form;
        [DataMember]
        public string name_1;
        [DataMember]
        public string last_name_ukr_1;
        [DataMember]
        public string name_ukr_1;
        [DataMember]
        public int chair_number;
        [DataMember]
        public int chair_number_1;
    }
}

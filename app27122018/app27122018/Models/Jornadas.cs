//------------------------------------------------------------------------------
// <auto-generated>
//     Ce code a été généré à partir d'un modèle.
//
//     Des modifications manuelles apportées à ce fichier peuvent conduire à un comportement inattendu de votre application.
//     Les modifications manuelles apportées à ce fichier sont remplacées si le code est régénéré.
// </auto-generated>
//------------------------------------------------------------------------------

namespace app27122018.Models
{
    using System;
    using System.Collections.Generic;
    
    public partial class Jornadas
    {
        public int JornadaID { get; set; }
        public Nullable<System.DateTime> dia { get; set; }
        public Nullable<int> EmpleadoID { get; set; }
        public string Horario { get; set; }
    
        public virtual Empleados Empleados { get; set; }
    }
}
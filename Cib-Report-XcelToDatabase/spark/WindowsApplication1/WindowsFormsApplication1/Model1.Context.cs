﻿//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated from a template.
//
//     Manual changes to this file may cause unexpected behavior in your application.
//     Manual changes to this file will be overwritten if the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace WindowsFormsApplication1
{
    using System;
    using System.Data.Entity;
    using System.Data.Entity.Infrastructure;
    
    public partial class CIBEntities : DbContext
    {
        public CIBEntities()
            : base("name=CIBEntities")
        {
        }
    
        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            throw new UnintentionalCodeFirstException();
        }
    
        public virtual DbSet<Com_i> Com_i { get; set; }
        public virtual DbSet<D_Contract_History> D_Contract_History { get; set; }
        public virtual DbSet<d_Other_sub_linked> d_Other_sub_linked { get; set; }
        public virtual DbSet<DETAILS_OF_INSTALL_Faca> DETAILS_OF_INSTALL_Faca { get; set; }
        public virtual DbSet<Flag> Flags { get; set; }
        public virtual DbSet<I_ADDRESS> I_ADDRESS { get; set; }
        public virtual DbSet<I_INQUIRED> I_INQUIRED { get; set; }
        public virtual DbSet<IMaster> IMasters { get; set; }
        public virtual DbSet<owner_list> owner_list { get; set; }
        public virtual DbSet<PROP_CONCERN> PROP_CONCERN { get; set; }
        public virtual DbSet<REQUESTED_CONTRACT> REQUESTED_CONTRACT { get; set; }
        public virtual DbSet<Sub__INFO> Sub__INFO { get; set; }
        public virtual DbSet<SUM_OF_FACILITY_S_AS_BOR> SUM_OF_FACILITY_S_AS_BOR { get; set; }
        public virtual DbSet<SUM_OF_FUNDED_FACILI_AS_BOR> SUM_OF_FUNDED_FACILI_AS_BOR { get; set; }
        public virtual DbSet<SUM_OF_NON_FUNDED_FACILI_AS_BOR> SUM_OF_NON_FUNDED_FACILI_AS_BOR { get; set; }
        public virtual DbSet<sysdiagram> sysdiagrams { get; set; }
    }
}

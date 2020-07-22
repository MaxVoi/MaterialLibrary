namespace FriendOrganizer.DataAccess.Migrations
{
    using Model;
    using System;
    using System.Data.Entity;
    using System.Data.Entity.Migrations;
    using System.Linq;

    internal sealed class Configuration : DbMigrationsConfiguration<FriendOrganizer.DataAccess.FriendOrganizerDbContext>
    {
        public Configuration()
        {
            AutomaticMigrationsEnabled = false;
        }

        protected override void Seed(FriendOrganizer.DataAccess.FriendOrganizerDbContext context)
        {
            context.Friends.AddOrUpdate(
                f=>f.FirstName,
                new Friend { FirstName = "Thomas", LastName = "Edison" },
                new Friend { FirstName = "Andras", LastName = "Boehler" },
                new Friend { FirstName = "Julia", LastName = "Huber" },
                new Friend { FirstName = "Chrissi", LastName = "Egin" });
        }
    }
}
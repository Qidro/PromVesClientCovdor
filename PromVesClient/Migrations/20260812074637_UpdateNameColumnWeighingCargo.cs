using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace PromVesClient.Migrations
{
    /// <inheritdoc />
    public partial class UpdateNameColumnWeighingCargo : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.RenameColumn(
                name: "Сargo",
                table: "Weighings",
                newName: "Cargo");
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.RenameColumn(
                name: "Cargo",
                table: "Weighings",
                newName: "Сargo");
        }
    }
}

using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace PromVesClient.Migrations
{
    /// <inheritdoc />
    public partial class UpdateNameColumnWeighing : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.RenameColumn(
                name: "Сonsignee",
                table: "Weighings",
                newName: "Consignee");
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.RenameColumn(
                name: "Consignee",
                table: "Weighings",
                newName: "Сonsignee");
        }
    }
}

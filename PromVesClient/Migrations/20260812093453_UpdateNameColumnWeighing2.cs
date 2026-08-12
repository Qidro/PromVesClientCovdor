using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace PromVesClient.Migrations
{
    /// <inheritdoc />
    public partial class UpdateNameColumnWeighing2 : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.RenameColumn(
                name: "InvoiceDataTime",
                table: "Weighings",
                newName: "InvoiceDateTime");
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.RenameColumn(
                name: "InvoiceDateTime",
                table: "Weighings",
                newName: "InvoiceDataTime");
        }
    }
}

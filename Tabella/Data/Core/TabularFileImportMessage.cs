using MessageKit.Configuration.Interfaces;
using MessageKit.Utility.Builders;
using MessageKit.Data.Core;
using MessageKit.Enums;

namespace Tabella.Data.Core;
 
 /// <summary>
 /// Message class for tabular file import-related messages.
 /// </summary>
 public class TabularFileImportMessage : Message
 {
     /// <summary>
     /// The translation key for the message.
     /// </summary>
     public new string TranslationKey { get; private set; }
     
     /// <summary>
     /// The name of the sheet where the message is relevant.
     /// </summary>
     public string? SheetName { get; }
     
     /// <summary>
     /// The index of the row where the message is relevant.
     /// </summary>
     public int? RowIndex { get; }
 
     /// <summary>
     /// The name of the column where the message is relevant.
     /// </summary>
     public string? ColumnName { get; }
     
     /// <summary>
     /// The value of the cell where the message is relevant.
     /// </summary>
     public string? CellValue { get; }
     
     /// <summary>
     /// The column e.g. "A0" where the message is relevant.
     /// </summary>
     public string? Column { get; }
 
     /// <summary>
     /// Initializes a new instance of the <see cref="TabularFileImportMessage"/> class.
     /// </summary>
     /// <param name="messageBuilder"></param>
     /// <param name="messageConfiguration"></param>
     /// <param name="sender"></param>
     /// <param name="sheetName"></param>
     /// <param name="rowIndex"></param>
     /// <param name="cellData"></param>
     /// <param name="level"></param>
     /// <param name="sentAt"></param>
     /// <param name="readAt"></param>
     public TabularFileImportMessage(
         MessageBuilder messageBuilder, 
         IMessageConfiguration messageConfiguration,
         string sender = "System", 
         string? sheetName = null,
         int? rowIndex = null, 
         TabularFileCellData? cellData = null,
         MessageInformationTypeEnum level = MessageInformationTypeEnum.Error, 
         DateTime sentAt = default, 
         DateTime? readAt = null) : base(messageBuilder, messageConfiguration, sender, level, sentAt, readAt)
     {
         TranslationKey = messageConfiguration.ResolveTranslationKey(null);
         
         SheetName = sheetName;
         RowIndex = rowIndex;
         
         if (cellData is not null)
         {
             ColumnName = cellData.ColumnName;
             CellValue = cellData.CellValue;
             Column = cellData.ColumnIndex.HasValue ? ToSheetedFileColumnName(cellData.ColumnIndex.Value) : null;
         }
         
         if(SheetName is null && RowIndex is null && ColumnName is null && CellValue is null && Column is null)
         {
             TranslationKey = "messages.tabular_file_import_base_no_extra_data";
         }
         
         PopulateCorePlaceholders();
     }
     
     private void PopulateCorePlaceholders()
     {
         CorePlaceholders[nameof(SheetName)] = SheetName ?? string.Empty;
         CorePlaceholders[nameof(RowIndex)] = RowIndex?.ToString() ?? string.Empty;
         CorePlaceholders[nameof(ColumnName)] = ColumnName ?? string.Empty;
         CorePlaceholders[nameof(CellValue)] = CellValue ?? string.Empty;
         CorePlaceholders[nameof(Column)] = Column ?? string.Empty;
     }
     
     /// <summary>
     /// Serves as the default hash function.
     /// </summary>
     /// <returns></returns>
     public override int GetHashCode()
     {
         return HashCode.Combine(
             MessageTranslationKey,
             GetDictionaryHashCode(MessagePlaceholders),
             InformationType,
             SheetName,
             RowIndex,
             ColumnName,
             CellValue,
             Column
         );
     }
     
     private static string ToSheetedFileColumnName(int columnIndex)
     {
         if (columnIndex < 1)
             return string.Empty; 

         string columnName = string.Empty;
         int dividend = columnIndex;

         while (dividend > 0)
         {
             int modulo = (dividend - 1) % 26;
             columnName = Convert.ToChar('A' + modulo) + columnName;
             dividend = (dividend - modulo) / 26;
         }

         return columnName;
     }
 }
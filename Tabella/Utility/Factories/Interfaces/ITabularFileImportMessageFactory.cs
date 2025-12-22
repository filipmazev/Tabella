using MessageKit.Utility.Builders;
using Tabella.Data.Core;
using MessageKit.Enums;

namespace Tabella.Utility.Factories.Interfaces;

/// <summary>
/// Factory interface for creating message instances.
/// </summary>
public interface ITabularFileImportMessageFactory
{
    /// <summary>
    /// Creates a new TabularFileImportMessage instance using the provided parameters.
    /// </summary>
    /// <param name="messageBuilder"></param>
    /// <param name="sender"></param>
    /// <param name="sheetName"></param>
    /// <param name="rowIndex"></param>
    /// <param name="cellData"></param>
    /// <param name="level"></param>
    /// <param name="sentAt"></param>
    /// <param name="readAt"></param>
    /// <returns></returns>
    TabularFileImportMessage CreateMessage(
        MessageBuilder messageBuilder,
        string sender = "System", 
        string? sheetName = null,
        int? rowIndex = null, 
        TabularFileCellData? cellData = null,
        MessageInformationTypeEnum level = MessageInformationTypeEnum.Error, 
        DateTime sentAt = default, 
        DateTime? readAt = null);
}
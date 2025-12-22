using MessageKit.Configuration.Interfaces;
using Tabella.Utility.Factories.Interfaces;
using Microsoft.Extensions.Options;
using MessageKit.Utility.Builders;
using Tabella.Configuration;
using Tabella.Data.Core;
using MessageKit.Enums;

namespace Tabella.Utility.Factories;

/// <summary>
/// Factory class for creating Message instances.
/// </summary>
public class TabularFileImportMessageFactory(
    IMessageConfiguration messageConfig,
    IOptions<TabularFileImportMessageOptions> options)
    : ITabularFileImportMessageFactory
{
    private readonly TabularFileImportMessageOptions _options = options.Value;
    
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
    public TabularFileImportMessage CreateMessage(
        MessageBuilder messageBuilder,
        string? sender = null, 
        string? sheetName = null,
        int? rowIndex = null, 
        TabularFileCellData? cellData = null,
        MessageInformationTypeEnum level = MessageInformationTypeEnum.Error, 
        DateTime sentAt = default, 
        DateTime? readAt = null)
    {
        return new TabularFileImportMessage(
            messageBuilder: messageBuilder, 
            messageConfiguration: messageConfig, 
            sender: sender ?? _options.DefaultSender,
            sheetName: sheetName,
            rowIndex: rowIndex, 
            cellData: cellData, 
            level: level, 
            sentAt: sentAt,
            readAt: readAt);
    }
}
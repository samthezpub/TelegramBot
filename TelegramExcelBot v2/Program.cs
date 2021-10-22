using System;
using System.Threading;
using System.Threading.Tasks;
using Telegram.Bot;
using Telegram.Bot.Exceptions;
using Telegram.Bot.Extensions.Polling;
using Telegram.Bot.Types;
using Telegram.Bot.Types.Enums;
using Spire.Xls;
using System.Numerics;
using System.Reflection;
using System.Diagnostics;

var botClient = new TelegramBotClient("2048096323:AAF53G9L2fIMBP_aVHHnxz-0CkgvFDFW_D0");
//Creates workbook
Workbook workbook = new Workbook();




workbook.LoadFromFile(@"C:\Users\User\source\repos\TelegramExcelBot v2\TelegramExcelBot v2\bin\Debug\net5.0\Sample.xls");
//Gets first worksheet

Worksheet sheet = workbook.Worksheets[0];
string temp;
string tempstr;
string masOfCr = "Контрольные или самостоятельные по таким предметам:\n";
int pos = 4;
string[] masOfPredmets = new string[5] {" ","Укр.лит", "Информатика", "Украинский", "Математика"};
for (int i = 1; i <= 5; i++)
{
    temp = i.ToString();
    Console.WriteLine(temp);
    tempstr = "E" + temp;
    if (sheet.Range[tempstr].Text == "+")
    {
        tempstr = "A" + temp;
        masOfCr += "\t" + sheet.Range[tempstr].Text + "\n";

        Console.WriteLine(sheet.Range[tempstr].Text);
    }
}


var me = await botClient.GetMeAsync();
Console.WriteLine(
    $"Hello, World! I am user {me.Id} and my name is {me.FirstName}."
);

using var cts = new CancellationTokenSource();

// StartReceiving does not block the caller thread. Receiving is done on the ThreadPool.
botClient.StartReceiving(
    new DefaultUpdateHandler(HandleUpdateAsync, HandleErrorAsync),
    cts.Token);

Console.WriteLine($"Слушаю @{me.Username}");
Console.ReadLine();

// Send cancellation request to stop bot
cts.Cancel();

Task HandleErrorAsync(ITelegramBotClient botClient, Exception exception, CancellationToken cancellationToken)
{
    var ErrorMessage = exception switch
    {
        ApiRequestException apiRequestException => $"Telegram API Error:\n[{apiRequestException.ErrorCode}]\n{apiRequestException.Message}",
        _ => exception.ToString()
    };

    Console.WriteLine(ErrorMessage);
    return Task.CompletedTask;
}
async Task HandleUpdateAsync(ITelegramBotClient botClient, Update update, CancellationToken cancellationToken)
{
    if (update.Type != UpdateType.Message)
        return;
    if (update.Message.Type != MessageType.Text)
        return;

    var chatId = update.Message.Chat.Id;
    ////if (update.Message.Text[0] == '/' && update.Message.Text[1] != 'с')
    ////{
    ////    Console.WriteLine($"Получена команда '{update.Message.Text}' от {chatId}.");
    ////    await botClient.SendTextMessageAsync(chatId: chatId,
    ////        text: "Выполняю команду - " + update.Message.Text);
    ////}
    
    if (update.Message.Text == "/дз")
    {
        Console.Write("Получает дз: ");
        Console.WriteLine(update.Id);
        string Dz = "";
        for (int i = 2; i <= 5; i++)
        {

            Dz += sheet.Range["A" + i.ToString()].Text;
            Dz += ": ";
            Dz+= sheet.Range["B" + i.ToString()].Text;
            Dz += "\n";
        }
        

        await botClient.SendTextMessageAsync(
        chatId: chatId,
        
        text: "Выполняю команду - " + update.Message.Text + ":\n" + Dz
        ) ;
    }
    else if(update.Message.Text == "/start")
    {
        await botClient.SendTextMessageAsync(chatId: chatId, text:
            "/дз - домашка\n" +
            "/кр - контрольные\n" +
            "/ср - аналог /кр\n" +
            "/хелп - справка по командам\n" +
            "/сет - команда позволяющаяя помочь установить дз (/сет НазваниеПредмета ДЗ)\n");
    }
    else if(update.Message.Text == "/хелп" || update.Message.Text == "/help")
    {
        await botClient.SendTextMessageAsync(chatId: chatId, text:
            "/дз - домашка\n" +
            "/кр - контрольные\n" +
            "/ср - аналог /кр\n" +
            "/хелп - справка по командам\n" +
            "/сет - команда позволяющаяя помочь установить дз (/сет НазваниеПредмета ДЗ)\n");
    }
    else if (update.Message.Text == "/кр" || update.Message.Text == "/ср")
    {
        await botClient.SendTextMessageAsync(chatId: chatId,
            text: masOfCr
            );
    }
    
    else if (update.Message.Text.Contains("/сет"))
    {
        Console.Write("Послан запрос: ");
        Console.Write(update.Message.Text);
        Console.Write(" ,от ");
        Console.WriteLine(chatId);
        await botClient.SendTextMessageAsync(chatId: chatId, text: "Вы послали запрос с содержимым " + update.Message.Text);
        string message = update.Message.Text;
       
        
        string[] Redacted = message.Split(' ');
        Console.WriteLine(Redacted[1]);
        Console.WriteLine(Redacted[2]);

        for (int i = 2; i <= 5; i++)
        {
            if (sheet.Range["A" + i.ToString()].Text.Contains(Redacted[1]))
            {
                pos = i;
            }

        }
        //Process.Start(Assembly.GetEntryAssembly().Location);

        //Process.GetCurrentProcess().Kill();
        sheet.Range["B" + pos.ToString()].Text = Redacted[2];
    }
    else
{
    Console.WriteLine($"Получено сообщение '{update.Message.Text}' от {chatId}.");
    await botClient.SendTextMessageAsync(
    chatId: chatId,
    text: "такой команды нет эээээээээээээээээээ\nээээээээ\n/хелп для справки, алёёёёёё"
);
}
    workbook.SaveToFile("Sample.xls");
    try
    {
        System.Diagnostics.Process.Start(workbook.FileName);
    }
    catch { }
}

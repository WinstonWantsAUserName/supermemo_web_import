# supermemo_web_import
 AutoHotkey script for quick import from browser to SuperMemo.

Select text and press `Ctrl + Shift + Alt + A` in the browser (currently supporting Chrome, Firefox and Microsoft Edge) to import to SuperMemo. Press `Ctrl + Alt + A` for quick import without GUI.

For incremental web browsing (that is, not importing the whole article but instead highlight on the browser), press `Ctrl + Shift + Alt + B`. You need a highlighter plugin (eg, [Super Simple Highlighter](https://chromewebstore.google.com/detail/super-simple-highlighter/hhlhjgianpocpoppaiihmlpgcoehlhio)) for this to work, and you need to set the highlight shortcut to `Shift + Alt + H`. Video guide: https://www.youtube.com/watch?v=fUOiNeFtVBk

Please reload the script (via right clicking the AHK icon on the task bar) if something went wrong. This script depends on the [UIA library](https://github.com/Descolada/UIAutomation) and it might not work all the time (eg, not being able to retrieve current page's URL)--reloading mostly does the trick. Sometimes clicking in empty spaces in the web page fixes it too.


Features:

- format cleaned
- references included (title, link); at some websites (like YouTube) author and date are also included

To do:

- add explanation for video and "online element" import
- make a video about the above feature because it requires extra preparations in SuperMemo

You can support me here: https://ko-fi.com/winstonwolf or https://www.buymeacoffee.com/winstonwolf or https://www.paypal.com/paypalme/winstonwolfie

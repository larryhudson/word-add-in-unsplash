# Word add-in for Unsplash

This is a basic example add-in for Word, that allows you to search for stock images using Unsplash and insert them into your Word doc.

You can read how to build it here:
[Build a custom Word add-in to insert images from the Unsplash API](https://ghost.larryhudson.io/word-add-in-unsplash/)

To get this example working:
1. Clone this repository
2. Run npm install (need to be using Node v16 LTS)
3. Add your own Unsplash API key in `src/taskpane/components/App.js`
4. If on Mac, run `npm run dev-server` in one Terminal window and `npm run start` in another. If not on Mac, just run `npm run start`. A new Word window should pop up with 'Show Taskpane' in the Home tab of the ribbon.

You can find out more on Microsoft's official documentation site:
[Word add-ins documentation](https://docs.microsoft.com/en-us/office/dev/add-ins/word/)
# Skype For Business integration for your Motion Openchannel in Node.js

## Overview

The key to an effective customer service is focusing in the channels where your customers are. Know your customers and be accessible to them through the most important channels. It might be through Email, Chat, Facebook, Twitter or other social media platforms. It’s true that there are too many channels and you might think you need a lot of time and resource. Don’t worry! xCALLY Motion Open Channel enables you to develop your favorite channel in just few steps. You will be able to easily manage your customer interactions through any channel inside a single Omni channel Desktop interface. Let’s see how you can develop your favorite customer service channel.
xCALLY Motion provides all the necessary tools to receive and send messages. All you need to do is implement a simple web service to exchange messages between xCALLY Motion server and your favorite channel. You can use any programming language you prefer.

## Prerequisites

  * [Node.js](http://nodejs.org/)
  * [Git](http://git-scm.com/)


## Setting up the app
  * Set the open channel [Account on Motion](https://wiki.xcallymotion.com/display/XMD/Open+Channel)
  * Download the code `git clone https://github.com/xcally/openchannel-skype4business.git`
  * Please see `config.json` in the root folder to change the default application settings.
  * Run `npm install` at the root folder to download dependencies.
  * Run `node index.js` to start the application.

## Configuration

Please see `config.json` in the root folder if you want to change the default application settings.

```javascript
{
    "microsoft_app_id": "YOUR_MICROSOFT_APP_ID",
    "microsoft_app_password": "YOUR_MICROSOFT_APP_PASSWORD",
    "url": "http://YOUR_MOTION_DOMAIN/api/openchannel/accounts/OPENCHANNEL_INTEGRATION_ID/receive",
    "port": 3003
}
```

## Contribution & support

Want to contribute? Great! Just create a pull request and you're in - welcome!

## Troubleshooting


## Enjoy

Thank you for choosing XCALLY MOTION, one of the first Omni Channel solution integrated with AsteriskTM and the most innovative real time solutions available on the market.

For more information, please visit our website [www.xcallymotion.com](https://www.xcallymotion.com/)

---
title: "Async Python Handlers in AWS Lambda"
date: 2020-06-23
slug: async-python-handlers-in-aws-lambda
tags: [software development]
---

#### Context

I am creating my first AWS Lambda function in Python from scratch. My function contains async functions, so I declared my handler with the async keyword. But when I deployed and smoke tested in AWS, I received the following error:

```
{
"errorMessage": "Unable to marshal response: Object of type coroutine is not JSON serializable",
"errorType": "Runtime.MarshalError"
}
```

#### Search terms

aws lambda python async handler

#### Helpful link

[https://stackoverflow.com/questions/60455830/can-you-have-an-async-handler-in-lambda-python-3-6](https://stackoverflow.com/questions/60455830/can-you-have-an-async-handler-in-lambda-python-3-6)

#### Outcome

Apparently AWS Lambda does not support async handler functions in Python. So you are required to call your asynchronous code from a synchronous handler function using this line of code:

```
asyncio.get_event_loop().run_until_complete(your_async_handler())
```

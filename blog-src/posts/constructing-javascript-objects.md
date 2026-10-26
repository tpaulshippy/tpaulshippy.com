---
title: "Constructing Javascript Objects"
date: 2020-06-18
slug: constructing-javascript-objects
---

#### Context

I am building a simple model class in Typescript/Javascript and was wondering if there was a shorthand way to avoid having to explicitly set each property name based on the constructor arguments.

#### Search terms

javascript constructor properties shorthand

#### Helpful link

[https://maksimivanov.com/posts/typescript-constructor-shorthand/](https://maksimivanov.com/posts/typescript-constructor-shorthand/)

#### Outcome

I found exactly what I was looking for. This is a Typescript feature only, so it wouldn’t work in pure Javascript, but it is much easier to type:

```
class Something {
  constructor(public prop1: string, public prop2: string) {}
}
```

than to type:

```
class Something {
  public prop1: string;
  public prop2: string;
  constructor(prop1: string, prop2: string) {
    this.prop1 = prop1;
    this.prop2 = prop2;
  }
}
```

Both produce the same Javascript in the end.

---
title: "Building a Native iOS App in a Day"
date: 2026-10-13
slug: building-a-native-ios-app-in-a-day
---

I did a thing this past weekend for the first time in my life. I built a fully native iOS app and got it production ready in about 24 hours.

This was not React Native (how I built [syftlearning.app](https://syftlearning.app/) and [careerality.app](https://careerality.app/)). This was 100% Swift. See the source code [here](https://github.com/tpaulshippy/verseloop)  but I will be honest with you – I haven’t. That’s right, I have not looked at the code. I don’t have very much familiarity with Swift, and haven’t felt the need to learn it. After this experience, I definitely still don’t.

How did I build it? Well, let me give you the full story. Let’s start at 10:43pm last Friday night. I tapped on OpenCode in my iPhone which is set to use the Muse Spark 1.3 Contributor model and is running on my Mac M1. I typed up this prompt:

> Goal — Build a native ios app that replaces a major workflow of one of my top most used apps [screenshot of iOS Screen Time: Signal 1h33m, X 1h10m, Spotify 1h8m, …] with an innovative user experience that is fun but also productive and smart. 
> Finish condition — I have an app that I will want to use built for me and distributable to my iPhone se3. 
> Evidence — A new public github that describes the app and contains the full source code along with screen recordings of the app running in simulator. 
> Constraints — must not clone; must be better performing; server (if any) runnable on a free cloudflare worker; must not fully replace the existing app, just enable a major workflow in a new way. 
> Build in a subfolder. Go.

I got the idea for these elements of the prompt from [this tweet](https://x.com/dotta/status/2108665034226225425?s=20). Result: agent built a Spotify “dice DJ” concept first. I prompted with this:

> **Hate it. Pitch me your next 10 ideas**

It gave me a bunch of other stuff that was mostly not great, including this:

> **5. NYT Games → Mini Coach (for: Wordle/Connections warmup)** Workflow: getting better, not playing the actual puzzle. UX: 60-second drills — anagram tiles, letter-ladder, near-miss patterns. Never shows today’s answer. Smart: Spaced repetition on your weak letters. Performs: 100% offline, instant. Server: none.

A mini coach for warming up to play Wordle. Haha. No thanks.

But it also included this:

> 8. **Bible → Verse Loop (for: memorizing 1 verse/week)** Workflow: daily verse retention. UX: Fill-in-the-blank swipe game, 30s/day, streak flame. Not a reader. Smart: SM-2 spaced repetition on-device. Performs: offline KJV text bundled. Server: none.

I had been thinking for a while that I might want my own Bible memorization app. I switched from Fighter Verses to Versify about a year ago, and it works ok. But I added spaced repetition to [Syft Learning](https://syftlearning.app/) recently and so I was familiar with that concept and thought it might be helpful. So I prompted with this:

> I like 8. Currently I start my day using Versify for this but there are weaknesses — no real spaced repetition from my feedback; first-letter testing punishes accidental taps with no recourse; everything is device-local, no sync or export. I definitely need ESV — I only memorize that. Go do the finish condition, evidence, and constraints with this idea as the goal.

I submitted this prompt to my agent at 10:57pm Friday. At 11:13pm [this commit](https://github.com/tpaulshippy/verseloop/commit/484a75d1db7b28e299488a9cd1ffcfc5174140e1) was pushed to a brand new repo called “VerseLoop” on my GitHub. If you take a look at that commit you can see that it included screenshots of the app (made possible because my MacBook was set up with xcode and simulator already). So I didn’t even need to open my laptop, let alone install the app on my phone to know what had been built.

Based on looking at those screenshots alone, I followed up with this at 8:26am the next morning:

> Fix these issues –
> The main recall interface is wrong. It is not supposed to be saying it out loud based on seeing the first letters. It should be text inputs where the user types each first letter in one at a time and when they get it right it shows that word and when they get it wrong it vibrates, marks it as a mistake, and shows the word.
> The accidental tap feature needs to be on each incorrect letter.
> Here is a screenshot of the interface in verify. [image-1.png] I don’t think we need the red progress bar at the top, but showing a count of mistakes might be helpful.
> Dark mode needs to be supported and the recordings should be in dark mode.
> Fix and push to main.

I provided this screenshot from Versify (accidentally called it “verify” in order to give it an example.

![](/wp-content/uploads/sites/2/2026/10/image-1-576x1024.png)

Handling this took the agent quite a while (46 minutes) but then I had [this commit](https://github.com/tpaulshippy/verseloop/commit/74883ce3c5622aa69be74ea13432530d7482b847) which actually included the core workflow I was looking for.

Next, one more refinement:

> Don’t show the reference while testing. The reference should be the last thing you test. First letter of the book and then the full numeric reference (except the colon) should be required.
> Don’t show the number of words in the verse.
> Update all the screenshots to be in dark mode.

At this point, I liked what I was seeing and wanted to install it. So I asked my agent to get the app to TestFlight. I did have to go download a p8 file from App Store Connect to give it access, provide some IDs, and select my team in xcode. The agent was also not able to actually create the app for me, so I did that manually. The name “Verse Loop” was taken, so I went with “Verse Loop Daily.”

At about 2pm, I was able to install the app on my phone and start testing. Here is the next feedback prompt that I gave:

> Ok it works. I don’t like how the reference is tested though. It should be displayed under the verse just like the words in the verse. And it should be a seamless experience. For Psalm 23:1 I should be able to type “tlimsisnwp231” in one smooth session and see it flow in. Tighten up the user interface by removing only what is necessary. Reference the screenshot I gave you of versify

At this point, I worked with the agent to set up xcode cloud which seemed to be the most cost effective way to get a fairly short deploy/test cycle going.

By about 4pm, I had the agent create a server deployment to a CloudFlare worker so that the app could both provide backup mechanisms as well as connect to the ESV API.

At 5:41, I gave this feedback:

> Let’s cleanup this interface also. 
> ![](/wp-content/uploads/sites/2/2026/10/image.png)
> We don’t need:
> visible textbox with “first letter…” Just bring up the keyboard and take inputs without showing a textbox or cursor.
> the instruction label about typing the first letter of the next word (we can put this in a tutorial or help screen later if needed)
> show word should be a button with a help icon
> restart should be a button with a relevant icon

Here are a sampling of prompts from the rest of the evening as I finished getting it to MVP state:

> Do we support multi verse passages? If not, add that.

> Add a way for the user to paste in json to import as an alternative to providing a file.

> On longer passages that require scrolling, the app should auto scroll as the user types letters.

> On the loop page (the main screen) let’s not show the start of the verse. The user needs to recall this.

> ![](/wp-content/uploads/sites/2/2026/10/image-2-576x1024.png)
> Why is the top item not showing as due? When I tap on it, it does.

As you can see from the screenshot, the app was starting to really come together. At about 9:20pm, the original chat thread with the agent started throwing errors, so I switched to a fresh thread with no context. By that point, I was polishing the import/export and backup mechanisms. I was able to try out the import with my own verses so that I could practice with real passages.

At about 10pm Saturday evening, I asked my agent to prepare the app for submission to the app store. It gave me all the material I needed including screenshots, descriptions, etc. After a bit of cleanup, I pointed my [Hark agent](https://hark.com/) at the repo and asked it to submit the app to the app store for me.

![](/wp-content/uploads/sites/2/2026/10/image-3-1007x1024.png)

As you can see from the screenshot, Hark was able to upload the screenshots and fill out most of the fields. It had trouble with the Description, because of some special character I think, so I copied/pasted that manually from [here](https://github.com/tpaulshippy/verseloop/blob/main/appstore/metadata.md#description-4000-chars).

The app was submitted at 11:49pm Saturday evening, so about 25 hours after I started building it.

I did try it out both Sunday and this (Monday) morning to practice my passages in Romans 11, and added a few notification features and fixed a bug today. They will come out in the next update.

At 2:44pm today, Apple notified me that they had approved the app for the app store and I published it at 4:12pm.

Go [check it out](https://t.co/CqDt3YNLfq) and let me know what you think! Reach out with any questions you might have about my process.

<div align="center">

## Debugging VB Applications


</div>

### Description

The purpose of this article is to show how to use the VB IDE to accomplish debugging. And why use ActiveX components in your projects.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[CodeDoctor](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/codedoctor.md)
**Level**          |Intermediate
**User Rating**    |3.0 (6 globes from 2 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Debugging and Error Handling](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/debugging-and-error-handling__1-26.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/codedoctor-debugging-vb-applications__1-54375/archive/master.zip)





### Source Code

```
Debugging and Component Development
By CodeDoctor
Part 1 – Debugging Explained
The purpose of this article is to show how to use the VB IDE to accomplish good sound debugging. I’ve been a member here for quite sometime, and this is my first post, as I have not noticed any information for the beginner to Intermediate level programmers that show debugging techniques that are a must for any VB programmer. This only covers VB3-VB6. VB.Net will soon have most of these abilities in future releases, so stay tuned for a revised article covering VB.Net debugging in the near future.
This article will also cover semi-advanced topics such as creating and debugging ActiveX DLL's. You can expect at the end of this article to be able to debug any VB application.
Visual Basic has come a long way over the years. One of the primary reasons is due to its rapid application environment, or IDE (Integrated Development Environment). The Code editor of VB6 allows developers to step through their code at many different levels. To give you an example of what this means, consider the following: (Place a single command button on a new Standard EXE project form)
When clicking on the Command button, the Command1_Click procedure fires, resulting in a msgbox showing. So you may be saying, “what is the big deal? Everyone knows that!”
Yes, most everyone should know how to do this, however, when you need to stop the code, you can do this by setting breakpoints in code as shown below:
As you can see, the msgbox statement is now highlighted in red. There are three ways to accomplish this, the first, is to simply click on the gray margin (Circled in Red). The second, you can goto the Debug menu, and choose “Toggle Breakpoint”. Or, the easiest way, is to just hit F9 on your Keyboard.
Now, when you hit F5 to run the application and you click on the Command1 command button, the msgbox does not display, the execution of the application stops before showing the message box. As shown below:
This breakpoint is very helpful to debugging, and also is used to learn the execution of how a program is operating.
Several things can be done while in “Break Mode”. The main thing is that you can check the state of your application by using the Immediate Window, available from the View Menu / Immediate Window, or by hitting CTRL-G on your keyboard. To check the variable sMsg in the immediate window, simply type in Debug.Print sMsg directly in the Immediate Window, and you’ll see the value print out “This is a test”. You can also execute other built in functions and methods while in breakmode. To try this, simply type in Debug.Print ASC(“A”). You should see the number 65, which is the Ascii Value of the Capital Letter “A”.
Also, if your project had modules, you can even call your own functions and procedures here while in break mode too.
Part 2 – ActiveX Components and why use them
To give a better example, lets look at some more advanced ways of using the IDE’s debugger. This next section will show you how to create an ActiveX DLL, and how to debug it.
First, you may be asking, “why create an ActiveX DLL, if you can simply just make your own procedures in modules?”
The reason is simple, however, there are a few areas you may need to understand before understanding it fully.
The reason why ActiveX DLL’s are better than just code modules, is “Encapsulation”. This sounds to most like another Techno-Babble word. But it has great merit. To give you a better example, consider that when a person starts an automobile, most follow these steps;
1.	Put key into Ignition
2.	Turn Key forward while applying gas
3.	Release turn pressure from Ignition
If you notice, the person starting the car does not need to know how the start operates, or have to know anything at all about the wiring connected to the ignition, instead, they just turn the key.
This is Encapsulation, all of the inner workings of vehicles sub-systems are abstracted from the operator. The operator only needs the key to activate the many supporting sub systems that make the vehicle operate.
Lets take this and create an object model
Car
	Transmission
	Engine
		Starter
			Ignition
‘The VB Code might look something like this
Dim oCar as New Car
Dim oCarKey as New Key
Dim oIgnition as Ignition
Set oIgnition = oCar.Engine.Starter.Ignition
oCarKey.Owner = “CodeDoctor”
Set oIgnition.Key = oCarKey
If oCarKey.IsValid = true then
	Do Until oCar.Engine.IsRunning = true
		oIgnition.Start
		if oIgnition.Attempts > 10 then
			‘// Need a Tow Truck
			Exit Do
		end if
	Loop
Else
	Msgbox “Invalid Car Key”
End if
This example shows that each Car has an Engine, each Engine has a Starter, and each Starter has an Ignition. Encapsulation is achieved in this model by simply calling the oIgnition.Start method. This method would be quite complex, as it must connect to even more objects or sub-systems, such as the electrical, battery, the engines other sub-systems such as the fuel injectors, crank, pistons etc. The user only needs to know that the Ignition has a method named “Start”, and the user must always call this method, that effectively calls other sub-systems, and their methods.
So with this example, you should be able to see the importance of using ActiveX DLL’s, as the benefits can out-weigh any standard code modules. As the next example explains.
Say that you are creating a Race Track, and you need to access many cars at one time.
Dim oRaceCars as New Cars
Dim oRaceCar as Car
For each oRaceCar in oRaceCars
	oRaceCar.Engine.Starter.Ignition.Start
next
The above example shows that you can start ALL the cars by simply calling the same method on each object. This means the state of each object is “Abstracted” or “Hidden” from the User, and each object has its own “State”. This is much more efficient than using code modules, while maintaining the state of global variables would be very difficult.
Part 3 – Creating an ActiveX DLL component
Coming Soon
```


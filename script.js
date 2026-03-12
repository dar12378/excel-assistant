const userPrompt=document.getElementById("userPrompt")

const generateBtn=document.getElementById("generateBtn")

const multiBtn=document.getElementById("multiBtn")

const clearBtn=document.getElementById("clearBtn")

const errorBox=document.getElementById("errorBox")

const singleResult=document.getElementById("singleResult")

const multiResult=document.getElementById("multiResult")

const formulaText=document.getElementById("formulaText")

const explanationText=document.getElementById("explanationText")

const exampleText=document.getElementById("exampleText")

const tipsList=document.getElementById("tipsList")

const formulaGrid=document.getElementById("formulaGrid")

const copyBtn=document.getElementById("copyBtn")

const historyList=document.getElementById("historyList")

const examplesList=document.getElementById("examplesList")

const clearHistoryBtn=document.getElementById("clearHistoryBtn")

const defaultColumnInput=document.getElementById("defaultColumn")



function excelFormula(name,args){

return "=" + name + "(" + args + ")"

}



function buildFormula(prompt){

let column=defaultColumnInput.value||"A"

prompt=prompt.toLowerCase()



if(prompt.includes("סכום")){

return{

formula:excelFormula("SUM",column+":"+column),

explanation:"מחבר את כל הערכים בעמודה",

example:"סכום כל הערכים",

tips:["אפשר להשתמש גם בטווח כמו A1:A10"]

}

}



if(prompt.includes("ממוצע")){

return{

formula:excelFormula("AVERAGE",column+":"+column),

explanation:"מחשב ממוצע",

example:"ממוצע ציונים",

tips:["מתעלם מתאים ריקים"]

}

}



if(prompt.includes("ספור")){

return{

formula:excelFormula("COUNTIF",column+":"+column+",\"Yes\""),

explanation:"סופר כמה פעמים מופיע Yes",

example:"ספירת סטטוס",

tips:["אפשר לשנות Yes"]

}

}



if(prompt.includes("גדול")){

return{

formula:excelFormula("IF","A2>100,\"כן\",\"לא\""),

explanation:"בודק אם גדול מ100",

example:"בדיקה פשוטה",

tips:["אפשר לשנות מספר"]

}

}



return{

formula:excelFormula("SUM","A:A"),

explanation:"נוסחת ברירת מחדל",

example:"סכום",

tips:["נסי לכתוב יותר ברור"]

}

}



generateBtn.onclick=()=>{

let prompt=userPrompt.value

if(!prompt){

errorBox.innerText="יש לכתוב בקשה"

errorBox.classList.remove("hidden")

return

}



let result=buildFormula(prompt)



formulaText.innerText=result.formula

explanationText.innerText=result.explanation

exampleText.innerText=result.example



tipsList.innerHTML=""

result.tips.forEach(t=>{

let li=document.createElement("li")

li.innerText=t

tipsList.appendChild(li)

})



singleResult.classList.remove("hidden")



saveHistory(prompt)

}



multiBtn.onclick=()=>{

let column=defaultColumnInput.value||"A"



formulaGrid.innerHTML=""



let formulas=[

excelFormula("SUM",column+":"+column),

excelFormula("AVERAGE",column+":"+column),

excelFormula("MAX",column+":"+column),

excelFormula("MIN",column+":"+column)

]



formulas.forEach(f=>{

let card=document.createElement("div")

card.className="example-item"

card.innerHTML="<pre>"+f+"</pre>"

formulaGrid.appendChild(card)

})



multiResult.classList.remove("hidden")

}



copyBtn.onclick=()=>{

navigator.clipboard.writeText(formulaText.innerText)

}



clearBtn.onclick=()=>{

userPrompt.value=""

singleResult.classList.add("hidden")

multiResult.classList.add("hidden")

}



function saveHistory(text){

let history=JSON.parse(localStorage.getItem("excelHistory")||"[]")

history.unshift(text)

history=history.slice(0,5)

localStorage.setItem("excelHistory",JSON.stringify(history))

renderHistory()

}



function renderHistory(){

let history=JSON.parse(localStorage.getItem("excelHistory")||"[]")

historyList.innerHTML=""



history.forEach(h=>{

let div=document.createElement("div")

div.className="history-item"

div.innerText=h

div.onclick=()=>userPrompt.value=h

historyList.appendChild(div)

})

}



clearHistoryBtn.onclick=()=>{

localStorage.removeItem("excelHistory")

renderHistory()

}



const examples=[

"חבר את כל הערכים בעמודה B",

"חשב ממוצע של עמודה F",

"ספור כמה פעמים Approved מופיע בעמודה C",

"בדוק אם A2 גדול מ100"

]



examples.forEach(e=>{

let div=document.createElement("div")

div.className="example-item"

div.innerText=e

div.onclick=()=>userPrompt.value=e

examplesList.appendChild(div)

})



renderHistory()

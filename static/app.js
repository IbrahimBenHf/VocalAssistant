class Chatbox{
    constructor(){
        this.args ={
            openbutton: document.querySelector('.chatbox__button'),
            chatbox: document.querySelector('.chatbox__support'),
            sendbutton: document.querySelector('.send__button'),
        }
        this.state  =false;
        this.messages=[];

    }

    display(){
        const {openbutton, chatbox, sendbutton} =this.args;
        openbutton.addEventListener('click',()=>this.toggleState(chatbox))
        sendbutton.addEventListener('click',()=>this.onSendButton(chatbox))
        const  node = chatbox.querySelector('input');
        node.addEventListener("keyup",({key}) =>{
            if(key==="Enter"){
                this.onSendButton(chatbox)
            }
            }
        )


        let final_transcript;
        final_transcript = '';
        const button = document.getElementById('button');
        button.addEventListener('click', () => {
            if (button.style['animation-name'] === 'flash') {
                recognition.stop();
                button.style['animation-name'] = 'none';
                button.innerText = 'Start';
                let msg = {name: "User", message: final_transcript};
                this.messages.push(msg);
                this.updateChatText(chatbox)
                fetch($SCRIPT_ROOT + '/predict', {
                    method:'POST',
                    body:JSON.stringify({message:final_transcript, language:this.getLanguage(),mail:this.getMail(),question:this.getLastQuestion()}),
                    mode:'cors',
                    headers:{
                        'Content-Type':'application/json'
                    },
                })
                    .then(r=>r.json())
                    .then(r=>{
                        let msg2 = {name : "Bot", message:r.answer};
                        this.messages.push(msg2);
                        this.speak(msg2.message);
                        this.updateChatText(chatbox)
                        textField.value = ''
                    }).catch((error)=>{
                    console.error('Error:',error);
                    this.updateChatText(chatbox)
                    textField.value =''
                });
                final_transcript = '';
            } else {
                button.style['animation-name'] = 'flash';
                button.innerText = 'Stop';
                recognition.start();
            }
        })

        const button1 = document.getElementById('button1');
        button1.addEventListener('click', () => {
            var e = document.getElementById("language");
            var strUser = e.value;
            console.log(strUser);

            var x = document.getElementById("mail");
            var xx = x.value;
            console.log(xx);
            window.speechSynthesis.speak(new SpeechSynthesisUtterance('Bonjour'));
        })

        const recognition = new webkitSpeechRecognition();
        recognition.lang = 'en'
        recognition.continuous = true;
        recognition.interimResults = true;
        recognition.onresult = function (event) {

            var interim_transcript = '';

            for (var i = event.resultIndex; i < event.results.length; ++i) {
                if (event.results[i].isFinal) {
                    final_transcript += event.results[i][0].transcript;
                } else {
                    interim_transcript += event.results[i][0].transcript;
                }
            }
        }

    }
    toggleState(chatbox){
        this.state = !this.state;
        if (this.state){
            chatbox.classList.add('chatbox--active')
        } else {
            chatbox.classList.remove('chatbox--active')
        }
    }

    speak(text){
        let utter = new SpeechSynthesisUtterance();
        utter.lang = 'en';
        utter.text = text;
        utter.volume = 0.5;
        window.speechSynthesis.speak(utter);
    }
    onSendButton(chatbox){
        var textField = chatbox.querySelector('input');
        let text1 =textField.value
        if (text1===""){
            return;
        }

        let msg1 ={name : "User", message:text1}
        this.messages.push(msg1);



        fetch($SCRIPT_ROOT + '/predict', {
            method:'POST',
            body:JSON.stringify({message:text1, language:this.getLanguage(),mail:this.getMail(),question:this.getLastQuestion()}),
            mode:'cors',
            headers:{
                'Content-Type':'application/json'
            },
        })
            .then(r=>r.json())
            .then(r=>{
                let msg2 = {name : "Bot", message:r.answer};
                this.messages.push(msg2);
                this.updateChatText(chatbox)
                textField.value = ''
            }).catch((error)=>{
                console.error('Error:',error);
                this.updateChatText(chatbox)
            textField.value =''
        });
    }

    updateChatText(chatbox){
        var html = '';
        this.messages.slice().reverse().forEach(function (item,index) {
            if (item.name  ==="Bot"){
                html+='<div class="messages__item messages__item--visitor">'+item.message+'</div>'
            }
            else{
                html+='<div class="messages__item messages__item--operator">'+item.message+'</div>'
            }

        });
        const chatmessage = chatbox.querySelector('.chatbox__messages');
        chatmessage.innerHTML = html;
    }

    getMail(){
        var e = document.getElementById("mail");
        var mail = e.value;
        return mail;
    }
    getLanguage(){
        var e = document.getElementById("language");
        var language = e.value;
        return language;
    }
    getLastQuestion(){
        if (this.messages.length === 1){
          return "no question";
        }
        return this.messages[this.messages.length - 2].message
    }

    wait(ms){
        var start = new Date().getTime();
        var end = start;
        while(end < start + ms) {
            end = new Date().getTime();
        }
    }
}

const chatbox = new Chatbox();

chatbox.display();
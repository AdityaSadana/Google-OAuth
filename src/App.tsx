import './App.css';
import { useState } from 'react';
import { useGoogleLogin } from '@react-oauth/google';
import axios from 'axios'
import { OPENAI_API_KEY } from './secrets';

interface IEmail {
  from: string,
  subject: string,
  date: number,
  body: string
}

function App() {

  const [accessToken, setAccessToken] = useState<string>()
  const [emails, setEmails] = useState<IEmail[]>([])
  const [selectedEmail, setSelectedEmail] = useState<IEmail>()
  const [predictedLabel, setPredictedLabel] = useState<string>()
  const [generatedReply, setGeneratedReply] = useState<string>()

  const login = useGoogleLogin({
    onSuccess: (tokenResponse) => {setAccessToken(tokenResponse.access_token); console.log(tokenResponse)},
    onError: (errorResponse) => console.log(errorResponse),
    scope: "https://mail.google.com/"
  });

  function compareEmails(email1: IEmail, email2: IEmail) {
    return email2.date - email1.date
  }

  const fetchLatestEmails = () => {

    axios.get(`https://gmail.googleapis.com/gmail/v1/users/me/messages?access_token=${accessToken}&maxResults=5&labelIds=INBOX`)
    .then((res) => {
      const topFive = res.data.messages
      console.log(topFive)

      let allEmails: IEmail[] = []
      Promise.all(
        topFive.map((m: any) => {
          return (axios.get(`https://gmail.googleapis.com/gmail/v1/users/me/messages/${m.id}?access_token=${accessToken}`)
          .then(messageResponse => {
            const from = messageResponse.data.payload.headers.find((h: any) => h.name === "From").value
            const subject = messageResponse.data.payload.headers.find((h: any) => h.name === "Subject").value
            const date = new Date(messageResponse.data.payload.headers.find((h: any) => h.name === "Date").value).getTime()
            const body = messageResponse.data.snippet
            allEmails.push({from, subject, date, body})
          })
          .catch(err => {
            alert(err)
            console.error(err)
          })
        )})
      )
      .then(() => {
        console.log("Fetched all messages")
        allEmails.sort(compareEmails)
        setEmails(allEmails)
      })
      .catch(err => {
        alert(err)
        console.error(err)
      })
    })
    .catch(err => {
      alert(err)
      console.error(err)
    })
  }

  const predictLabel = () => {
    const chat = {
      model: "gpt-3.5-turbo",
      messages: [{
        role: "system",
        content: "You are a helpful assistant that drafts emails."
      },
      {
        role: "user",
        content: selectedEmail!.body
      },
      {
        role: "user",
        content: "Predict Label (Interested, Not-Interested, Needs more information) for above message"
      }],
      temperature: 0.2
    }

    axios.post("https://api.openai.com/v1/chat/completions", chat, {headers: {
      "Authorization": `Bearer ${OPENAI_API_KEY}`
    }})
    .then(res => {
      /*
      Sample Output res.data => 
        {
          "id": "chatcmpl-9MhmFjJVNIaw18mEe1iqhuPv4aTeF",
          "object": "chat.completion",
          "created": 1715198523,
          "model": "gpt-3.5-turbo-0125",
          "choices": [
            {
              "index": 0,
              "message": {
                "role": "assistant",
                "content": "Needs more information"
              },
              "logprobs": null,
              "finish_reason": "stop"
            }
          ],
          "usage": {
            "prompt_tokens": 47,
            "completion_tokens": 3,
            "total_tokens": 50
          },
          "system_fingerprint": null
        }
      */
      setPredictedLabel(res.data.choices[0].message.content)
    })
    .catch(err => {
      alert(err)
      console.error(err)
    })
  }

  const sendMessage = async () => {
    const message = createMessage(selectedEmail!.from, selectedEmail!.subject, generatedReply!);

    try {
      const res = await axios.post(
        'https://www.googleapis.com/gmail/v1/users/me/messages/send',
        {
          raw: message,
        },
        {
          headers: {
            Authorization: `Bearer ${accessToken}`,
          },
        }
      );
      console.log('Email sent successfully:', res.data);
      alert('Email sent successfully!');
    } catch (error) {
      console.error('Error sending email:', error);
      alert('Failed to send email.');
    }
  };

  const createMessage = (to: string, subject: string, body: string) => {
    const emailLines = [
      'Content-Type: text/plain; charset="UTF-8"\n',
      `To: ${to}\n`,
      `Subject: ${subject}\n\n`,
      `${body}`,
    ];
    return btoa(emailLines.join('')).replace(/\+/g, '-').replace(/\//g, '_');
  };

  const generateReply = () => {
    const chat = {
      model: "gpt-3.5-turbo",
      messages: [{
        role: "system",
        content: "You are a helpful assistant that drafts emails."
      }, {
        role: "user",
        content: selectedEmail!.body
      }, {
        role: "user",
        content: "Generate Reply for the above email without subject."
      }],
      temperature: 0.2
    }

    axios.post("https://api.openai.com/v1/chat/completions", chat, {headers: {
      "Authorization": `Bearer ${OPENAI_API_KEY}`
    }})
    .then(res => {
      setGeneratedReply(res.data.choices[0].message.content)
    })
    .catch(err => {
      alert(err)
      console.error(err)
    })
  }

  return (
      <div className="App">
        {accessToken ?
          <>
            {selectedEmail ?
              <div className='flex flex-col w-full'>
                <button className='w-[50px]' onClick={() => {setSelectedEmail(undefined); setPredictedLabel(undefined); setGeneratedReply(undefined)}}>Back</button>
                <div className='flex'>
                  <p className='w-[80%] my-[12px]'>{selectedEmail.body}</p>
                  {predictedLabel ?
                    <p className='m-auto'>{predictedLabel}</p>
                    :
                    <button className='m-auto w-[200px]' onClick={predictLabel}>Predict Label</button>
                  }
                </div>
                {generatedReply ?
                  <div className='flex flex-col align-center'>
                    <p className='font-semibold'>{generatedReply}</p>
                    <button className='w-[200px] mx-auto my-[10px]' onClick={sendMessage}>Send Message</button>
                  </div>
                  :
                  <button className='w-[200px]' onClick={generateReply}>Generate Reply</button>
                }
              </div>
              :
              <div className='flex flex-col w-full'>
                <div className='mx-auto text-center'>
                  <h2 className='font-bold text-2xl'>You are successfully signed in !</h2>
                  <button className='m-[10px]' onClick={fetchLatestEmails}>Fetch Emails</button>
                </div>
                
                <div className='flex flex-col'>
                  {emails.length > 0 &&
                    <div className='border-bottom flex justify-between'>
                      <p className='border-right w-[25%] text-center font-bold text-lg'>From</p>
                      <p className='border-right w-[25%] text-center font-bold text-lg'>Subject</p>
                      <p className='border-right w-[40%] text-center font-bold text-lg'>Body</p>
                      <p className='w-[8%] font-bold text-lg'>Option</p>
                    </div>
                  }
                  {emails.map(e => {
                    return (
                      <div className='border-bottom flex justify-between text-center py-[32px]'>
                        <p className='border-right w-[25%]'>{e.from}</p>
                        <p className='border-right w-[25%]'>{e.subject}</p>
                        <p className='border-right w-[40%]'>{e.body}</p>
                        <button className='w-[8%] h-[40px]' onClick={() => {setSelectedEmail(e)}}>Open</button>
                      </div>
                    )
                  })}
                </div>
              </div>
            }
          </>
          :
          <div>
            <button onClick={() => login()}>Login</button>
          </div>
        }
      </div>
  );
}

export default App;

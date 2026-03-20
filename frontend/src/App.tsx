import { useState } from 'react'
import { Message, sendMessageStream } from './api/chat'
import ChatWindow from './components/ChatWindow'
import InputBar from './components/InputBar'

export default function App() {
  const [messages, setMessages] = useState<Message[]>([])
  const [isStreaming, setIsStreaming] = useState(false)
  const [streamingText, setStreamingText] = useState('')

  const handleSend = (text: string) => {
    const newMessages: Message[] = [...messages, { role: 'user', content: text }]
    setMessages(newMessages)
    setIsStreaming(true)
    setStreamingText('')

    let accumulated = ''
    sendMessageStream(
      newMessages,
      (chunk) => {
        accumulated += chunk
        setStreamingText(accumulated)
      },
      () => {
        setMessages((prev) => [...prev, { role: 'assistant', content: accumulated }])
        setIsStreaming(false)
        setStreamingText('')
      }
    )
  }

  const handleReset = () => {
    setMessages([])
    setStreamingText('')
    setIsStreaming(false)
  }

  return (
    <div className="flex flex-col h-screen bg-gray-900 text-gray-100">
      <header className="flex items-center justify-between px-6 py-3 border-b border-gray-700">
        <h1 className="text-lg font-semibold">Osstem AI Chat</h1>
        <button
          onClick={handleReset}
          className="text-xs text-gray-400 hover:text-gray-200 transition-colors"
        >
          대화 초기화
        </button>
      </header>
      <ChatWindow messages={messages} isStreaming={isStreaming} streamingText={streamingText} />
      <InputBar onSend={handleSend} disabled={isStreaming} />
    </div>
  )
}

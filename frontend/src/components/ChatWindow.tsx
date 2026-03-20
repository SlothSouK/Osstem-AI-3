import { useEffect, useRef } from 'react'
import { Message } from '../api/chat'
import MessageBubble from './MessageBubble'

interface Props {
  messages: Message[]
  isStreaming: boolean
  streamingText: string
}

export default function ChatWindow({ messages, isStreaming, streamingText }: Props) {
  const bottomRef = useRef<HTMLDivElement>(null)

  useEffect(() => {
    bottomRef.current?.scrollIntoView({ behavior: 'smooth' })
  }, [messages, streamingText])

  return (
    <div className="flex-1 overflow-y-auto px-4 py-4">
      {messages.length === 0 && (
        <div className="flex h-full items-center justify-center text-gray-500">
          <p>안녕하세요! 무엇이든 물어보세요.</p>
        </div>
      )}
      {messages.map((msg, i) => (
        <MessageBubble key={i} message={msg} />
      ))}
      {isStreaming && (
        <div className="flex justify-start mb-3">
          <div className="max-w-[75%] rounded-2xl rounded-bl-sm px-4 py-2 text-sm bg-gray-700 text-gray-100 whitespace-pre-wrap break-words">
            {streamingText}
            <span className="inline-block w-1 h-4 bg-gray-400 animate-pulse ml-0.5 align-middle" />
          </div>
        </div>
      )}
      <div ref={bottomRef} />
    </div>
  )
}

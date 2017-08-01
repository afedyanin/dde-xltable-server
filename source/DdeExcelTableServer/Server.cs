namespace DdeExcelTableServer
{
    using System;
    using System.Timers;
    using NDde.Server;

    public class Server : DdeServer
    {
        private readonly Timer timer = new Timer();

        public Action<string> OnBeforeConnectAction;
        public Action<string, string> OnAfterConnectAction;
        public Action<string, string> OnDisconnectAction;
        public Action<string, string, byte[], int> OnPokeAction;

        public Server(string serviceName) : base(serviceName)
        {
            this.timer.Elapsed += this.OnTimerElapsed;
            this.timer.Interval = 1000;
            this.timer.SynchronizingObject = this.Context;
            this.timer.AutoReset = false;
        }

        public override void Register()
        {
            base.Register();
            this.timer.Start();
        }

        public override void Unregister()
        {
            this.timer.Stop();
            base.Unregister();
        }

        protected override bool OnBeforeConnect(string topic)
        {
            this.OnBeforeConnectAction?.Invoke(topic);
            return true;
        }

        protected override void OnAfterConnect(DdeConversation conversation)
        {
            this.OnAfterConnectAction?.Invoke(conversation.Service, conversation.Topic);
        }

        protected override void OnDisconnect(DdeConversation conversation)
        {
            this.OnDisconnectAction?.Invoke(conversation.Service, conversation.Topic);
        }

        protected override PokeResult OnPoke(DdeConversation conversation, string item, byte[] data, int format)
        {
            this.OnPokeAction?.Invoke(conversation.Topic, item, data, format);
            return PokeResult.Processed;
        }

        private void OnTimerElapsed(object sender, ElapsedEventArgs args)
        {
            // Advise all topic name and item name pairs.
            this.Advise("*", "*");
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                this.timer.Dispose();
            }

            base.Dispose(disposing);
        }
    }
}

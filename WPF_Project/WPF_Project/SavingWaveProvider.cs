using NAudio.Wave;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WPF_Project
{
    internal class SavingWaveProvider : IWaveProvider
    {
        private readonly IWaveProvider sourceWaveProvider;
        private readonly WaveFileWriter writer;
        private bool isWriterDisposed;
        public WaveFormat WaveFormat { get { return sourceWaveProvider.WaveFormat; } }

        public SavingWaveProvider(IWaveProvider sourceWaveProvider, string wavFilePath)
        {
            this.sourceWaveProvider = sourceWaveProvider;
            writer = new WaveFileWriter(wavFilePath, sourceWaveProvider.WaveFormat);
        }

        public int Read(byte[] buffer, int offset, int count)
        {
            var read = sourceWaveProvider.Read(buffer, offset, count);
            if (count > 0 && !isWriterDisposed)
            {
                writer.Write(buffer, offset, read);
            }
            if (count == 0)
            {
                Dispose(); // auto-dispose in case users forget
            }
            return read;
        }
        public void Dispose()
        {
            if (!isWriterDisposed)
            {
                isWriterDisposed = true;
                writer.Dispose();
            }
        }
    }
}

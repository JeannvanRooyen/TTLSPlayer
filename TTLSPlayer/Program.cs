using Microsoft.VisualBasic;
using NAudio.Midi;
using System.Diagnostics;
using System.Linq.Expressions;
using System.Reflection;
using System.Speech.Synthesis;
using System.Windows.Forms;

namespace TTLSPlayer
{
    public class SpeechSynth : IDisposable
    {
        public SpeechSynthesizer Synthesizer = new SpeechSynthesizer();

        public void ChooseVoice(int rate)
        {
            try
            {
                int voice = -1;
                string choice = "";
                while (voice < 0)
                {
                    Console.Clear();
                    Console.ForegroundColor = ConsoleColor.White;
                    Console.WriteLine("Available voices:");
                    Console.WriteLine("=================");
                    for (int i = 0; i < Synthesizer.GetInstalledVoices().Count; i++)
                    {
                        var v = Synthesizer.GetInstalledVoices()[i];
                        Console.WriteLine($"{i + 1}: {v.VoiceInfo.Name} : {v.VoiceInfo.Gender}");
                    }
                    Console.WriteLine("Please enter desired voice (Blank for default): ");
                    choice = Console.ReadLine() ?? "1";

                    voice = int.Parse(choice) - 1;
                }

                Synthesizer.SelectVoice(Synthesizer.GetInstalledVoices()[voice].VoiceInfo.Name);
                Synthesizer.Volume = 100;
                Synthesizer.Rate = rate;
            }
            catch (Exception err)
            {
                Console.WriteLine($"Error: {err.Message}");
            }
        }

        public void SpeakAsync(string lyrics)
        {
            Synthesizer.SpeakAsync(lyrics);
        }

        public void Speak(string lyrics)
        {
            Synthesizer.Speak(lyrics);
        }

        public void SpeakWordSSML(string word, int pitch)
        {
            try
            {
                string pitchString = pitch > 0 ? "+" + pitch.ToString() : pitch.ToString();
                string ssml = $"<speak version='1.0' xml:lang=\"en-US\" xmlns='http://www.w3.org/2001/10/synthesis'><prosody pitch='{pitchString}%'>{word}</prosody></speak>";

                Synthesizer.SpeakSsml(ssml);
            }
            catch (Exception err)
            {
                Console.WriteLine($"SPEAK SSML ERR: {err.Message}");
            }
        }

        public void Dispose()
        {
            Synthesizer.Dispose();
        }
    }

    public class NotePlayer : IDisposable
    {
        MidiOut midiOut;

        public NotePlayer()
        {
            midiOut = new MidiOut(0);
            createCMajorScale();
        }

        public void CloseMidi()
        {
            Dispose();
        }

        public List<NoteEntity> CMajorScale { get; set; } = new List<NoteEntity>();

        private void createCMajorScale()
        {
            CMajorScale.Add(new NoteEntity(4, 'c', 60,0));
            CMajorScale.Add(new NoteEntity(4, 'd', 62,12));
            CMajorScale.Add(new NoteEntity(4, 'e', 64,25));
            CMajorScale.Add(new NoteEntity(4, 'f', 65, 30));
            CMajorScale.Add(new NoteEntity(4, 'g', 67, 48));
            CMajorScale.Add(new NoteEntity(4, 'a', 69, 60));
            CMajorScale.Add(new NoteEntity(4, 'b', 71, 75));
        }

        public void PlayNote(char note, int duration)
        {
            if (!CMajorScale.Any(o => o.Note == note))
            {
                Console.WriteLine($"ERR: Note '{note}' not found");
            }

            var n = CMajorScale.First(o => o.Note == note).Midi;

            if (n > 0)
            {
                midiOut.Send(MidiMessage.StartNote(n, 100, 1).RawData);
                Thread.Sleep(duration);
                midiOut.Send(MidiMessage.StopNote(n, 0, 1).RawData);
            }
            else
            {
                throw new Exception("Note not found");
            }
        }

        public void PlayNoteAsync(char note, int duration)
        {
            Task.Run(() => PlayNote(note, duration));
        }

        public void Dispose()
        {
            midiOut.Dispose();
        }
    }

    public class NoteEntity
    {
        public int Octave { get; set; } = 4;

        public char Note { get; set; } = ' ';

        public int Midi { get; set; } = 60;

        public int Pitch { get; set; } = 0;

        public NoteEntity(int octave, char note, int midi, int pitch)
        {
            Octave = octave;
            Note = note;
            Midi = midi;
            Pitch = pitch;
        }

        public int DoubleOctave()
        {
            return Midi + 12;
        }

        public int HalfOctave()
        {
            return Midi - 12;
        }
    }

    internal class Program
    {
        public static int versesSpoken = 0;

        static void Main(string[] args)
        {
            try
            {
                var songContent = "";
                string[] lyrics;
                Console.ForegroundColor = ConsoleColor.Cyan;
                Console.WriteLine("Read from Song File? [y/n] ");
                if (Console.ReadKey().Key == ConsoleKey.Y)
                {
                    songContent = File.ReadAllText("Songs.txt");
                    lyrics = songContent.Split("\n".ToCharArray());
                }
                else
                {
                    Console.BackgroundColor = ConsoleColor.Black;
                    Console.ForegroundColor = ConsoleColor.DarkYellow;
                    Console.WriteLine("Please enter music to \"sing\" ;) - lines 1-6");
                    Console.WriteLine("Tip: Leave line empty for auto-gen i.e. 'Just Press Enter :)'");
                    Console.WriteLine("Line 1: ");
                    var line1 = Console.ReadLine();
                    if (string.IsNullOrEmpty(line1))
                    {
                        line1 = "[c]twinkle [g]twinkle [a]little [g]star";
                        Console.ForegroundColor = ConsoleColor.Magenta;
                        Console.WriteLine(line1);
                        Console.ForegroundColor = ConsoleColor.DarkYellow;
                    }
                    Console.WriteLine("Line 2: ");
                    var line2 = Console.ReadLine();
                    if (string.IsNullOrEmpty(line2))
                    {
                        line2 = "[f]how [f]I [e]wonder [d]what [d]you [c]are";
                        Console.ForegroundColor = ConsoleColor.Magenta;
                        Console.WriteLine(line2);
                        Console.ForegroundColor = ConsoleColor.DarkYellow;
                    }
                    Console.WriteLine("Line 3: ");
                    var line3 = Console.ReadLine();
                    if (string.IsNullOrEmpty(line3))
                    {
                        line3 = "[g]up [f]above [f]the [e]cloud [e]so [d]high";
                        Console.ForegroundColor = ConsoleColor.Magenta;
                        Console.WriteLine(line3);
                        Console.ForegroundColor = ConsoleColor.DarkYellow;
                    }
                    Console.WriteLine("Line 4: ");
                    var line4 = Console.ReadLine();
                    if (string.IsNullOrEmpty(line4))
                    {
                        line4 = "[g]like [g]a [f]diamond [e]in [e]the [d]sky";
                        Console.ForegroundColor = ConsoleColor.Magenta;
                        Console.WriteLine(line4);
                        Console.ForegroundColor = ConsoleColor.DarkYellow;
                    }
                    Console.WriteLine("Line 5: ");
                    var line5 = Console.ReadLine();
                    if (string.IsNullOrEmpty(line5))
                    {
                        line5 = "[c]twinkle [g]twinkle [a]little [g]star";
                        Console.ForegroundColor = ConsoleColor.Magenta;
                        Console.WriteLine(line5);
                        Console.ForegroundColor = ConsoleColor.DarkYellow;
                    }
                    Console.WriteLine("Line 6: ");
                    var line6 = Console.ReadLine();
                    if (string.IsNullOrEmpty(line6))
                    {
                        line6 = "[f]how [f]I [e]wonder [d]what [d]you [c]are";
                        Console.ForegroundColor = ConsoleColor.Magenta;
                        Console.WriteLine(line6);
                        Console.ForegroundColor = ConsoleColor.DarkYellow;
                    }

                    Console.Clear();

                    lyrics = new string[] { line1, line2, line3, line4, line5, line6 };
                }

                Console.Clear();

                var voice = new SpeechSynth();
                voice.ChooseVoice(5);

                NotePlayer notePlayer = new NotePlayer();

                for (int i = 0; i < lyrics.Length; i++)
                {
                    if (string.IsNullOrEmpty(lyrics[i]))
                    {
                        Thread.Sleep(1000);
                        continue;
                    }

                    var music = "";
                   

                    for (int j = 0; j < lyrics[i].Length - 1; j++)
                    {
                        if (lyrics[i][j] == '[')
                        {
                            music += lyrics[i][j + 1];
                            music += "1";
                        }
                        else if (lyrics[i][j + 1] == ']')
                        {
                            j++;
                        }
                    }

                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine($"Verse {i}");
                    Console.ForegroundColor = ConsoleColor.White;
                    Console.WriteLine("--------------");
                    Console.WriteLine(lyrics[i]);

                    var pureLyrics = "";

                    for (int j = 0; j < lyrics[i].Length; j++)
                    {
                        if (lyrics[i][j] == '[')
                        {
                            j = j + 2;
                            continue;
                        }
                        else
                        {
                            pureLyrics += lyrics[i][j];
                        }
                    }
                    Console.ForegroundColor = ConsoleColor.White;
                    Console.WriteLine("--------------");
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"Music: {music}");
                    Console.ForegroundColor = ConsoleColor.White;
                    Console.WriteLine("--------------");
                    Console.ForegroundColor = ConsoleColor.Cyan;
                    Console.WriteLine($"Verse: {pureLyrics}");

                    var words = pureLyrics.Split(" ".ToCharArray());
                   
                    try
                    {
                        for (int j = 0; j < words.Length; j++)
                        {
                            notePlayer.PlayNoteAsync(music.Replace("1","")[j],100);
                            voice.SpeakWordSSML(words[j], 20);
                        }
                    }
                    catch (Exception errSmall)
                    {
                        Console.ForegroundColor = ConsoleColor.DarkRed;
                        Console.WriteLine($"Error: {errSmall.Message}");
                    }

                    Thread.Sleep(500);
                }

                voice.Dispose();
                notePlayer.CloseMidi();
            }
            catch (Exception err)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"Error: {err.Message}");
            }
            Console.WriteLine("\nDone. Press any key to exit.");
            Console.ReadKey();
        }
    }
}

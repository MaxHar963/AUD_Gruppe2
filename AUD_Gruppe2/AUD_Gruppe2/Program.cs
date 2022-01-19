List<string> listImages = new List<string>();
string[] files;

Console.WriteLine("Path to Picture-Folder:");
string path = Console.ReadLine();
//Console.WriteLine(path);
Console.WriteLine(" ");

if (path != null && path != "")
{
    files = Directory.GetFiles(path);

    foreach (var file in files)
    {
        listImages.Add(file);
        //Console.WriteLine("Path: " + path);
        //Console.WriteLine("Filename: " + file);
        //Console.WriteLine(" ");
    }

    foreach (var file in listImages)
    {
        Console.WriteLine(file.Split("\\").Last());
        Console.WriteLine(" ");
    }
}

Console.ReadKey();
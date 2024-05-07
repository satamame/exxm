namespace Args;

public class ComArgs
{
    string[] Args { get; set; }

    public string Mode { get; set; } = "";
    public bool Clean { get; set; }

    public ComArgs(string[] args)
    {
        this.Args = args;
    }

    public void Validate()
    {
        string mode = "";
        bool clean = false;

        foreach(var arg in this.Args)
        {
            if (mode != "" && arg == $"--{mode}")
            {
                throw new Exception($"{arg} が複数回指定されています。");
            }

            if (arg == "--version")
            {
                if (mode != "")
                {
                    throw new Exception($"--{mode} と --version は同時に指定できません。");
                }
                mode = "version";
            }
            else if (arg == "--help")
            {
                if (mode != "")
                {
                    throw new Exception($"--{mode} と --help は同時に指定できません。");
                }
                mode = "help";
            }
            else if (arg == "--from-excel")
            {
                if (mode != "")
                {
                    throw new Exception($"--{mode} と --from-excel は同時に指定できません。");
                }
                mode = "from-excel";
            }
            else if (arg == "--to-excel")
            {
                if (mode != "")
                {
                    throw new Exception($"--{mode} と --to-excel は同時に指定できません。");
                }
                mode = "to-excel";
            }
            else if (arg == "--clean")
            {
                if (clean)
                {
                    throw new Exception("--clean が複数回指定されています。");
                }
                clean = true;
            }
            else
            {
                throw new Exception($"Invalid argument: {arg}");
            }
        }

        if (mode == "")
        {
            throw new Exception("--from-excel または --to-excel を指定してください。");
        }

        this.Mode = mode;
        this.Clean = clean;
    }
}

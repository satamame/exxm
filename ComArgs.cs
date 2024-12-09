namespace Args;

public class ComArgs
{
    string[] Args { get; set; }

    public string Target { get; set; } = "";
    public string Mode { get; set; } = "";
    public bool Clean { get; set; } = false;

    public ComArgs(string[] args)
    {
        this.Args = args;
    }

    public void Validate()
    {
        string target = "";
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
                    throw new Exception($"--{mode} と {arg} は同時に指定できません。");
                }
                mode = "version";
            }
            else if (arg == "--help")
            {
                if (mode != "")
                {
                    throw new Exception($"--{mode} と {arg} は同時に指定できません。");
                }
                mode = "help";
            }
            else if (arg == "--from-xl")
            {
                if (mode != "")
                {
                    throw new Exception($"--{mode} と {arg} は同時に指定できません。");
                }
                mode = "from-xl";
            }
            else if (arg == "--to-xl")
            {
                if (mode != "")
                {
                    throw new Exception($"--{mode} と {arg} は同時に指定できません。");
                }
                mode = "to-xl";
            }
            else if (arg == "--clean")
            {
                // TODO: Implement --clean
                throw new Exception("--clean は今後のバージョンで実装予定です。");

                //if (clean)
                //{
                //    throw new Exception("--clean が複数回指定されています。");
                //}
                //clean = true;
            }
            else
            {
                if (target != "")
                {
                    throw new Exception($"対象が複数指定されています。: {target}, {arg}");
                }
                target = arg;
            }
        }

        if (mode == "")
        {
            throw new Exception("--from-xl または --to-xl を指定してください。");
        }

        this.Target = target;
        this.Mode = mode;
        this.Clean = clean;
    }
}

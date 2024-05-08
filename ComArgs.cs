namespace Args;

public class ComArgs
{
    string[] Args { get; set; }

    public string Mode { get; set; } = "";
    public bool Clean { get; set; } = false;

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
            else if (arg == "--from-excel")
            {
                if (mode != "")
                {
                    throw new Exception($"--{mode} と {arg} は同時に指定できません。");
                }
                mode = "from-excel";
            }
            else if (arg == "--from-xl")
            {
                if (mode != "")
                {
                    throw new Exception($"--{mode} と {arg} は同時に指定できません。");
                }
                mode = "from-xl";
            }
            else if (arg == "--to-excel")
            {
                if (mode != "")
                {
                    throw new Exception($"--{mode} と {arg} は同時に指定できません。");
                }
                mode = "to-excel";
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
                throw new Exception($"Invalid argument: {arg}");
            }
        }

        if (mode == "")
        {
            throw new Exception("--from-excel または --to-excel を指定してください。");
        }

        if (mode == "from-xl") mode = "from-excel";
        if (mode == "to-xl") mode = "to-excel";

        this.Mode = mode;
        this.Clean = clean;
    }
}

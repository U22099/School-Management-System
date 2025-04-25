{pkgs}: {
  channel = "stable-24.05";
  packages = [
    pkgs.nodejs_20
    pkgs.typescript
  ];
  idx.extensions = [
    
  ];
  idx.previews = {
    previews = {
      web = {
        command = [ 
          "npm" 
          "run" 
          "dev" 
          ];
        manager = "web";
      };
    };
  };
}
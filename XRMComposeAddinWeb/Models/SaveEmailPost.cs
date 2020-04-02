using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace XRMComposeAddinWeb.Models
{
    public class SaveEmailPost
    {
        public SaveEmailRequest fields { get; set; }
    }

  public class CreateUserInfo
  {
    public CreateUserDefaultConfigInfo fields { get; set; }
  }

  public class UpdateUserInfo
  {
    public UpdateUserDefaultConfigInfo fields { get; set; }
  }
}

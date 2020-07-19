      [cmdletbinding()]
      param(
        [Parameter(Position = 1)]
        [Object]$ComputerLocationHash
,
        [Parameter(Position = 0)]
        [String]$ComputerLocationFile
        )

$ComputerLocationHash = [Ordered]@{
  Department = [Ordered]@{
    InternalHash = @{
      Building = @{
        None = @{
          Room = @(
            0
          )
        }
      }
    }
    MCDO         = [Ordered]@{
      Building = [Ordered]@{
        AV29  = [Ordered]@{
          Room = @(
            8, 
            9, 
            10, 
            11, 
            12, 
            13, 
            14, 
            15, 
            16, 
            17, 
            18, 
            19, 
            20
          )
        }
        AV34  = [Ordered]@{
          Room = @(
            1, 
            6, 
            7, 
            8, 
            201, 
            202, 
            203, 
            204, 
            205
          )
        }
        ELC3  = [Ordered]@{
          Room = @(
            100, 
            101, 
            102, 
            103, 
            104, 
            105, 
            106, 
            107
          )
        }
        ELC31 = [Ordered]@{
          Room = @(
            1
          )
        }
        ELC32 = [Ordered]@{
          Room = @(
            1
          )
        }
        ELC33 = [Ordered]@{
          Room = @(
            1
          )
        }
        ELC34 = [Ordered]@{
          Room = @(
            1
          )
        }
        ELC35 = [Ordered]@{
          Room = @(
            1
          )
        }
        ELC36 = [Ordered]@{
          Room = @(
            1
          )
        }
      }
    }
    CA           = [Ordered]@{
      Building = [Ordered]@{
        AV29 = [Ordered]@{
          Room = @(
            1, 
            2, 
            3, 
            4, 
            5, 
            6, 
            7, 
            23, 
            24, 
            25, 
            26, 
            27, 
            28, 
            29, 
            30
          )
        }
        AV34 = [Ordered]@{
          Room = @(
            1, 
            2, 
            3, 
            200, 
            214
          )
        }
        AV44 = [Ordered]@{
          Room = @(
            1
          )
        }
        AV45 = [Ordered]@{
          Room = @(
            1
          )
        }
        AV46 = [Ordered]@{
          Room = @(
            1
          )
        }
        AV47 = [Ordered]@{
          Room = @(
            1, 
            2
          )
        }
        AV48 = [Ordered]@{
          Room = @(
            1, 
            2
          )
        }
      }
    }
    PRO          = [Ordered]@{
      Building = [Ordered]@{
        AV34 = [Ordered]@{
          Room = @(
            210, 
            211, 
            212, 
            213
          )
        }
        ELC4 = [Ordered]@{
          Room = @(
            1, 
            100, 
            101, 
            102, 
            103, 
            104, 
            105, 
            106, 
            107
          )
        }
      }
    }
    TJ           = [Ordered]@{
      Building = [Ordered]@{
        AV34 = [Ordered]@{
          Room = @(
            2, 
            3, 
            13, 
            11
          )
        }
        ELC2 = [Ordered]@{
          Room = @(
            1
          )
        }
      }
    }
  }
}

$ComputerLocationHash | ConvertTo-Json -Depth 5 | Out-File $ComputerLocationFile


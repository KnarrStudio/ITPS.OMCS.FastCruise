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
    Shipping         = [Ordered]@{
      Building = [Ordered]@{
        Warehouse1  = [Ordered]@{
          Room = @(
            8, 
20
          )
        }
        Warehouse2  = [Ordered]@{
          Room = @(
            1, 
            6
          )
        }
      }
    }
    Sales           = [Ordered]@{
      Building = [Ordered]@{
        TrumpTower = [Ordered]@{
          Room = @(
            101, 
            102, 
            103, 
            104, 
            105, 
            106
          )
        }

      }
    }

  }
}

$ComputerLocationHash | ConvertTo-Json -Depth 5 | Out-File $ComputerLocationFile


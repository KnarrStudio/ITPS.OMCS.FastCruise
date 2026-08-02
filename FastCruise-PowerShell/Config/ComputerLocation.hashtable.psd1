@{
    # =====================================================================
    # ComputerLocation.hashtable.psd1
    #
    # Edit this file to describe your Department > Building > Room
    # hierarchy, then run:
    #
    #   ConvertTo-LocationJson -SourceHashtableFile '.\ComputerLocation.hashtable.psd1' -DestinationJsonFile '.\ComputerLocation.json'
    #
    # to regenerate ComputerLocation.json. Editing a hashtable here is much
    # less error-prone than hand-editing the JSON file directly.
    # =====================================================================
    Department = @{
        Shipping = @{
            Building = @{
                Warehouse1 = @{
                    Room = @(8, 20)
                }
                Warehouse2 = @{
                    Room = @(1, 6)
                }
            }
        }
        Sales = @{
            Building = @{
                TrumpTower = @{
                    Room = @(101, 102, 103, 104, 105, 106)
                }
            }
        }
    }
}

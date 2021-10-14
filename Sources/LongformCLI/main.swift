//
//  main.swift
//  Longform Document Processor
//
//  Copyright (c) 2021 Eric Summers
//

import ArgumentParser
import Foundation
import Markdown

struct LongformCommand: ParsableCommand {

    static let configuration = CommandConfiguration(commandName: "longform", shouldDisplay: false, subcommands: [
        ]
    )
    
}

LongformCommand.main()

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
    enum Error: LocalizedError {
        case couldntDecodeInputAsUTF8

        var errorDescription: String? {
            switch self {
            case .couldntDecodeInputAsUTF8:
                return "Couldn't decode input as UTF-8."
            }
        }
    }

    static let configuration = CommandConfiguration(
        commandName: "longform",
        shouldDisplay: false,
        subcommands: [
        ]
    )
    
    @Argument(help: "Input file (default: standard input)")
    var inputFilePath: String?
    
    static func parseFile(at path: String, options: ParseOptions) throws -> (source: String, parsed: Document) {
        let data = try Data(contentsOf: URL(fileURLWithPath: path))
        guard let inputString = String(data: data, encoding: .utf8) else {
            throw Error.couldntDecodeInputAsUTF8
        }
        return (inputString, Document(parsing: inputString, options: options))
    }

    static func parseStandardInput(options: ParseOptions) throws -> (source: String, parsed: Document) {
        let stdinData: Data
        if #available(macOS 10.15.4, *) {
            stdinData = try FileHandle.standardInput.readToEnd() ?? Data()
        } else {
            stdinData = FileHandle.standardInput.readDataToEndOfFile()
        }
        guard let stdinString = String(data: stdinData, encoding: .utf8) else {
            throw Error.couldntDecodeInputAsUTF8
        }
        return (stdinString, Document(parsing: stdinString, options: options))
    }
    
    func run() throws {
        let source: String
        let document: Document
        let parseOptions: ParseOptions = [ .parseBlockDirectives ]
        if let inputFilePath = inputFilePath {
            (source, document) = try Self.parseFile(at: inputFilePath, options: parseOptions)
        } else {
            (source, document) = try Self.parseStandardInput(options: parseOptions)
        }
        
        // TODO
        let _ = source //
        print(document.debugDescription()) //
    }
    
}

LongformCommand.main()

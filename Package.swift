// swift-tools-version:5.5

import PackageDescription

let package = Package(
    name: "longform",
    platforms: [
        .macOS(.v10_15),
    ],
    products: [
        .library(
            name: "Longform",
            targets: ["Longform"]),
        .executable(
            name: "lf",
            targets: ["LongformCLI"]),
    ],
    dependencies: [
        .package(name: "swift-markdown", url: "https://github.com/apple/swift-markdown.git", .branch("main")),
        .package(url: "https://github.com/apple/swift-argument-parser.git", from: "0.4.4"),
        .package(url: "https://github.com/weichsel/ZIPFoundation.git", from: "0.9.14"),
    ],
    targets: [
        .executableTarget(
            name: "LongformCLI",
            dependencies: [
                "Longform",
                .product(name: "ArgumentParser", package: "swift-argument-parser"),
            ]
        ),
        .target(
            name: "Longform",
            dependencies: [
                .product(name: "Markdown", package: "swift-markdown"),
                .product(name: "ZIPFoundation", package: "ZIPFoundation"),
                "Wordprocessing",
            ]
        ),
        .target(
            name: "Wordprocessing",
            dependencies: [
                "OOXML",
            ]
        ),
        .target(
            name: "OOXML",
            dependencies: [
            ]
        ),
        .testTarget(
            name: "LongformTests",
            dependencies: [
                "Longform",
            ],
            resources: [
                .copy("Documents"),
            ]
        ),
    ]
)


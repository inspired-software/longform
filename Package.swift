// swift-tools-version:5.5

import PackageDescription

let package = Package(
    name: "longform",
    products: [
        .library(
            name: "Longform",
            targets: ["Longform"]),
        .executable(
            name: "longform",
            targets: ["LongformCLI"]),
    ],
    dependencies: [
        .package(url: "https://github.com/apple/swift-markdown.git", branch: "main"),
        .package(url: "https://github.com/apple/swift-argument-parser.git", from: "0.4.4"),
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
            ]
        ),
        .testTarget(
            name: "LongformTests",
            dependencies: [
                "Longform"
            ],
            resources: [
                .copy("Documents"),
            ]
        ),
    ]
)

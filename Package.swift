// swift-tools-version:5.5

import PackageDescription

let package = Package(
    name: "longform",
    platforms: [
        .macOS(.v11),
    ],
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
                "PagesScripting",
                .product(name: "Markdown", package: "swift-markdown"),
            ]
        ),
        .target(
            name: "PagesScripting",
            dependencies: [],
            cSettings: [
                .define("TARGET_OS_MACOS", to: "1", .when(platforms: [.macOS])),
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


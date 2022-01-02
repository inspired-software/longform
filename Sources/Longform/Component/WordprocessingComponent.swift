//
//  File.swift
//  
//
//  Created by Eric Summers on 1/2/22.
//

//import Foundation
//import Markdown
//
//public protocol WordprocessingComponent: Renderable {
//    /// The underlying component that should be used to render this component.
//    /// Can either be a `Node`, another `WordprocessingComponent`, or a group of components
//    /// created using the `WordprocessingGroup` type.
//    var body: WordprocessingComponent { get }
//}
//
///// A type that represents an empty, non-rendered component. It's typically
///// used within contexts where some kind of `Component` needs to be returned,
///// but when you don't actually want to render anything. Modifiers have no
///// affect on this component.
//public struct EmptyWordprocessingComponent: WordprocessingComponent {
//    /// Initialize an empty component.
//    public init() {}
//    public var body: WordprocessingComponent { Node<Any>.empty }
//}
//
//extension Markup {
//
//    static var empty: Markup {  }
//
//}

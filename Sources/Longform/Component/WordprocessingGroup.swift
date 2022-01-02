//
//  File.swift
//  
//
//  Created by Eric Summers on 1/2/22.
//

//import Foundation
//
///// Type used to define a group of components
/////
///// The `members` contained within a `WordprocessingGroup` act as one
///// unit when passed around, with the exception that any modifier
///// that is applied to a group will be applied to each member
///// individually. So, for example, applying the `class` modifier
///// to a group results in each element within that group getting
///// that class name assigned to it.
//public struct WordprocessingGroup: WordprocessingComponent {
//    /// The group's members. Will be rendered in order.
//    public var members: [WordprocessingComponent]
//    public var body: WordprocessingComponent { Node.components(members) }
//
//    /// Create a new group with a given set of member components.
//    /// - parameter members: The components that should be included
//    ///   within the group. Will be rendered in order.
//    public init(members: [WordprocessingComponent]) {
//        self.members = members
//    }
//}
//
//extension WordprocessingGroup: ComponentContainer {
//    public init(@ComponentBuilder content: () -> Self) {
//        self = content()
//    }
//}
//
//extension WordprocessingWordprocessingGroup: Sequence {
//    public func makeIterator() -> Array<WordprocessingComponent>.Iterator {
//        members.makeIterator()
//    }
//}

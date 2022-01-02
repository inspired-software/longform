//
//  File.swift
//  
//
//  Created by Eric Summers on 1/2/22.
//

import Foundation

///// Result builder used to combine all of the `WordprocessingComponent` expressions that appear
///// within a given attributed scope into a single `WordprocessingGroup`.
/////
///// You can annotate any function or closure with the `@WordprocessingBuilder` attribute
///// to have its contents be processed by this builder. Note that you never have to
///// call any of the methods defined within this type directly. Instead, the Swift
///// compiler will automatically map your expressions to calls into this builder type.
//@resultBuilder public enum WordprocessingBuilder {
//    /// Build a `ComponentGroup` from a list of components.
//    /// - parameter components: The components that should be included in the group.
//    public static func buildBlock(_ components: WordprocessingComponent...) -> WordprocessingGroup {
//        ComponentGroup(members: components)
//    }
//
//    /// Build a flattened `ComponentGroup` from an array of component groups.
//    /// - parameter groups: The component groups to flatten into a single group.
//    public static func buildArray(_ groups: [WordprocessingGroup]) -> WordprocessingGroup {
//        ComponentGroup(members: groups.flatMap { $0 })
//    }
//
//    /// Pick the first `ComponentGroup` within a conditional statement.
//    /// - parameter component: The component to pick.
//    public static func buildEither(first component: WordprocessingGroup) -> WordprocessingGroup {
//        component
//    }
//
//    /// Pick the second `ComponentGroup` within a conditional statement.
//    /// - parameter component: The component to pick.
//    public static func buildEither(second component: WordprocessingGroup) -> WordprocessingGroup {
//        component
//    }
//
//    /// Build a `ComponentGroup` from an optional group.
//    /// - parameter component: The optional to transform into a concrete group.
//    public static func buildOptional(_ component: WordprocessingGroup?) -> WordprocessingGroup {
//        component ?? WordprocessingGroup(members: [])
//    }
//}

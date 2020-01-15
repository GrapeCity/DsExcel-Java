//----------------------------------------------------------------------------------------
//	Copyright © 2007 - 2017 Tangible Software Solutions Inc.
//	This class can be used by anyone provided that the copyright notice remains intact.
//
//	This class is used to simulate the ability to pass arguments by reference in Java.
//----------------------------------------------------------------------------------------
/**
 *  This class is used to simulate interior_ptr&lt;T&gt; in Java. 
 */
public final class RefObject<T>
{
	public T value;
	/**
	  Wraps a ref (ByRef in Visual Basic) stack object.
	 */
	public RefObject(T refArg)
	{
		this.value = refArg;
	}
	/**
	  Wraps an out (&lt;Out&gt; ByRef in Visual Basic) parameter.
	 */
	public RefObject() {
		
	}
	
	@Override
	public String toString() {
		return this.value.toString();
	}
}
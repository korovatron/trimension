# Shared URL Restore Regression Checklist

Purpose: quick manual checks to confirm shared URL load restores all construction objects and labels, including derived-point dependencies.

How to run each case:
1. Build the scenario in the app.
2. Use Share to generate a URL.
3. Open the URL in a fresh tab (or after hard refresh).
4. Compare restored result with the original scene.

General pass criteria for every case:
1. Object count matches before and after share-load.
2. Labels match text and placement intent.
3. Points list includes expected derived points.
4. Diagram and Points sections are open after shared load.
5. Triangles, Segments, Angles, Planes, and Labels sections are collapsed after shared load.

## Case 1: Baseline Primitive Label

Setup:
1. Create one primitive.
2. Add one simple segment between two primitive points.
3. Add a label to that segment.

Expected:
1. Segment restores.
2. Label restores.

## Case 2: Midpoint on Primitive Edge

Setup:
1. Create one primitive.
2. Add midpoint to a primitive edge.
3. Add a segment from midpoint to another primitive point.
4. Add a label to that segment.

Expected:
1. Midpoint-derived point restores.
2. Segment that references midpoint restores.
3. Segment label restores.

## Case 3: Midpoint to Midpoint Segment

Setup:
1. Create one primitive.
2. Add two midpoint points on two primitive edges.
3. Add a segment between the two midpoint points.
4. Add a label to that midpoint-to-midpoint segment.

Expected:
1. Both midpoint-derived points restore.
2. Midpoint-to-midpoint segment restores.
3. Segment label restores.

## Case 4: Intersection to Midpoint Segment

Setup:
1. Create one primitive.
2. Add two segments that intersect to produce a derived intersection point.
3. Add midpoint on a primitive edge.
4. Add a segment from the derived intersection point to the derived midpoint point.
5. Add a label to that segment.

Expected:
1. Derived intersection point restores.
2. Derived midpoint point restores.
3. Segment from intersection to midpoint restores.
4. Segment label restores.

## Case 5: Mixed Construction With Additional Objects

Setup:
1. Create one primitive.
2. Add at least one triangle and one angle marker.
3. Add at least one midpoint and one intersection.
4. Add segments and labels that use both base and derived points.

Expected:
1. All geometry restores with no missing dependent objects.
2. All labels restore and remain attached to visible edges.
3. No restore alert is shown.

## Case 6: Visibility Persistence

Setup:
1. Build a mixed scene with segments, triangles, angles, and labels.
2. Hide one visible object from each type using the object visibility toggle.
3. Share and reopen.

Expected:
1. Hidden/visible state for each object is preserved.
2. Edge labels remain hidden when their backing edge is hidden.

## Case 7: Camera and Panel State

Setup:
1. Rotate and zoom camera to a distinctive view.
2. Share and reopen.

Expected:
1. Camera pose restores.
2. Control panel is open.
3. Diagram and Points sections are open.
4. Triangles, Segments, Angles, Planes, and Labels sections are collapsed.

## Quick Result Log

Date:
Tester:
Commit under test:

Case results:
1. Case 1: Pass or Fail
2. Case 2: Pass or Fail
3. Case 3: Pass or Fail
4. Case 4: Pass or Fail
5. Case 5: Pass or Fail
6. Case 6: Pass or Fail
7. Case 7: Pass or Fail

Notes:
